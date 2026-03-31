import streamlit as st

import pandas as pd

from datetime import datetime, timedelta, timezone

import time

import io

import pytz  # 시간대 설정을 위한 라이브러리

import streamlit.components.v1 as components # <-- 1. 이 라이브러리가 꼭 필요합니다!



# --- [세션 상태 초기화] ---

if 'df_raw' not in st.session_state: st.session_state.df_raw = None

if 'analyzed' not in st.session_state: st.session_state.analyzed = False

if 'p' not in st.session_state: st.session_state.p = {}

if 'add_order_dict' not in st.session_state: st.session_state.add_order_dict = {}

if 'upload_key' not in st.session_state: st.session_state.upload_key = 0



# --- [2. 새로고침 방지 경고창 스크립트] ---

# 이 코드가 실행되면 사용자가 F5를 누를 때 "정말 나갈 거냐"고 물어봅니다.

components.html(

    """

    <script>

    window.onbeforeunload = function() {

        return "데이터 분석 중입니다. 새로고침하면 작업 내용이 사라질 수 있습니다.";

    };

    </script>

    """,

    height=0, # 화면에는 안 보이게 높이를 0으로 설정

)



# 1. [환경 설정 - 한국 시간대 및 페이지 설정]

KST = pytz.timezone('Asia/Seoul') # 한국 시간대 정의

now = datetime.now(KST)          # 현재 한국 시간 가져오기



st.set_page_config(layout="wide", page_title="저스트원 재고관리 v4.0")



# --- [공통 함수: 구글 시트 연동] ---

def get_sheet():

    try:

        from oauth2client.service_account import ServiceAccountCredentials

        import gspread

        creds_dict = dict(st.secrets["gcp_service_account"])

        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)

        client = gspread.authorize(creds)

        return client.open_by_key("1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg")

    except:

        return None



# [필수 함수] 구글 시트 입고 기록 가져오기

def get_incoming_history():

    try:

        sheet = get_sheet() 

        ws = sheet.worksheet("입고기록")

        data = ws.get_all_records()

        if data:

            df_h = pd.DataFrame(data)

            df_h['상품명'] = df_h['상품명'].astype(str).str.strip()

            df_h['옵션'] = df_h['옵션'].astype(str).str.strip()

            summary = df_h.groupby(['상품명', '옵션'])['수량'].sum().reset_index()

            summary.rename(columns={'수량': '과거리오더 입고'}, inplace=True)

            return summary

        return pd.DataFrame(columns=['상품명', '옵션', '과거리오더 입고'])

    except:

        return pd.DataFrame(columns=['상품명', '옵션', '과거리오더 입고'])


# [추가 함수] 구글 시트1에서 기존 '리오더 수량' 컬럼만 가져오기
def sync_reorder_from_sheet(df_uploaded):
    try:
        # 1. 시트1 열기
        sh = get_sheet()
        ws = sh.worksheet("시트1")
        all_data = ws.get_all_records()
        
        if all_data:
            df_sheet = pd.DataFrame(all_data)
            
            # 2. 필수 컬럼 확인 (상품명, 옵션, 리오더 수량)
            if "상품명" in df_sheet.columns and "옵션" in df_sheet.columns and "리오더 수량" in df_sheet.columns:
                # 시트에서 필요한 것만 추출하여 정리
                df_ref = df_sheet[['상품명', '옵션', '리오더 수량']].copy()
                df_ref['상품명'] = df_ref['상품명'].astype(str).str.strip()
                df_ref['옵션'] = df_ref['옵션'].astype(str).str.strip()
                df_ref['리오더 수량'] = pd.to_numeric(df_ref['리오더 수량'], errors='coerce').fillna(0).astype(int)
                
                # 3. 현재 업로드된 데이터(df_uploaded)와 병합
                # 엑셀의 기존 '리오더 수량'은 무시하고 시트의 최신값을 가져옵니다.
                if "리오더 수량" in df_uploaded.columns:
                    df_uploaded = df_uploaded.drop(columns=["리오더 수량"])
                
                df_final = pd.merge(df_uploaded, df_ref, on=['상품명', '옵션'], how='left')
                df_final['리오더 수량'] = df_final['리오더 수량'].fillna(0).astype(int)
                
                return df_final
        return df_uploaded
    except Exception as e:
        st.error(f"⚠️ 시트에서 리오더 수량을 동기화하는 중 오류: {e}")
        return df_uploaded


# --- 시트 연결 테스트 모드 (필요할 때만 아래 줄들의 #을 지워서 사용하세요) ---

# st.sidebar.subheader("🔍 시트 연결 상태 점검")

# if st.sidebar.button("연결된 시트 탭 목록 확인하기"):

#     try:

#         sheet = get_sheet() # 기존에 만드신 시트 가져오는 함수

#         # 현재 구글 시트 파일 안에 있는 모든 탭(Worksheet) 이름을 가져옵니다.

#         worksheets = sheet.worksheets()

#         sheet_names = [s.title for s in worksheets]

#         

#         st.sidebar.success("✅ 시트 연결 성공!")

#         st.sidebar.write("**현재 발견된 탭 목록:**")

#         for name in sheet_names:

#             st.sidebar.code(name) # 탭 이름을 복사하기 좋게 코드로 출력

#             

#         # 필수 탭이 있는지 자동 체크

#         required = ["재고현황", "입고이력", "발주기록"]

#         for req in required:

#             if req in sheet_names:

#                 st.sidebar.write(f"✔️ `{req}`: 확인됨")

#             else:

#                 st.sidebar.error(f"❌ `{req}`: 탭을 찾을 수 없습니다!")

#                 

#     except Exception as e:

#         st.sidebar.error(f"❌ 시트 연결 실패: {e}")







# --- [세션 상태 초기화] ---

if 'df_raw' not in st.session_state: st.session_state.df_raw = None

if 'analyzed' not in st.session_state: st.session_state.analyzed = False

if 'p' not in st.session_state: st.session_state.p = {}

if 'add_order_dict' not in st.session_state: st.session_state.add_order_dict = {}



# [추가] 파일 업로드 위젯 리셋용 키 (이게 있어야 파일명이 지워집니다)

if 'upload_key' not in st.session_state: st.session_state.upload_key = 0



st.title("📦 저스트원 통합 재고 관리 v4.0")



tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 관리"])



with tab1:

    # --- 1단계: 데이터 업로드 ---

    st.subheader("📁 1단계: 데이터 업로드")

    

    # 파일 업로드 위젯 (초기화 시 파일명 삭제 기능 포함)

    up_file = st.file_uploader(

        "엑셀/CSV 파일 업로드", 

        type=['xlsx', 'xls', 'csv'], 

        key=f"up_file_{st.session_state.upload_key}"

    )

    

    # [🔄 화면 전체 초기화 버튼]

    if st.button("🔄 화면 전체 초기화", use_container_width=True):

        for key in list(st.session_state.keys()):

            if key != "upload_key": 

                del st.session_state[key]

        st.session_state.upload_key += 1

        st.session_state.analyzed = False 

        st.session_state.df_raw = None

        st.query_params.clear() 

        st.rerun()



  # 데이터 로드 로직 (이 블록 전체를 교체하세요)
    if up_file:
        if st.session_state.get('df_raw') is None:
            try:
                # 1. 먼저 업로드한 파일을 읽습니다.
                df = pd.read_csv(up_file) if up_file.name.endswith('.csv') else pd.read_excel(up_file)
                df.columns = df.columns.str.strip()
                
                # 2. ⭐ [중요] 여기서 시트의 '리오더 수량'을 가져와 합칩니다!
                # 아까 정의한 함수를 여기서 호출(사용)해야 데이터가 0이 안 됩니다.
                with st.spinner("🔄 구글 시트에서 기존 리오더 수량을 동기화 중..."):
                    df = sync_reorder_from_sheet(df)
                
                # 3. '리오더 수량' 컬럼이 없는 경우를 대비한 기본 처리
                if "리오더 수량" not in df.columns: 
                    df["리오더 수량"] = 0
                
                df = df.fillna("") 
                st.session_state.df_raw = df
                st.success("✅ 파일 업로드 및 시트 데이터 동기화 완료!")
                
            except Exception as e:
                st.error(f"파일 로드 오류: {e}")


   # --- 2~3단계: 매핑 및 분석 설정 ---
if st.session_state.get('df_raw') is not None:
    st.divider()
    
    # --- 2단계: 매핑 항목 ---
    st.subheader("📋 2단계: 매핑 항목")
    st.info("💡 좌측은 기본 정보, 우측은 수량 및 날짜 정보를 매칭해주세요.")
    cols = st.session_state.df_raw.columns.tolist()
    
    def auto_idx(keys, exclude_keys=None):
        for i, c in enumerate(cols):
            column_name = str(c)
            if exclude_keys and any(ek in column_name for ek in exclude_keys): continue
            if any(k in column_name for k in keys): return i
        return 0

    # 5개씩 2열로 배치 (좌: 기본정보 / 우: 수량 및 날짜)
    c_left, c_right = st.columns(2)
    
    with c_left:
        st.markdown("##### [ 기본 정보 ]")
        it = st.selectbox("📦 상품명", cols, index=auto_idx(['상품명']), key="sel_it")
        op = st.selectbox("🎨 옵션", cols, index=auto_idx(['옵션']), key="sel_op")
        vn = st.selectbox("🏭 공급처", cols, index=auto_idx(['공급처']), key="sel_vn")
        vi = st.selectbox("🆔 공급처 상품명", cols, index=auto_idx(['공급처상품명']), key="sel_vi")
        so = st.selectbox("🚫 품절 여부", cols, index=auto_idx(['품절']), key="sel_so")

    with c_right:
        st.markdown("##### [ 수량 및 날짜 ]")
        av = st.selectbox("✅ 가용재고", cols, index=auto_idx(['가용재고']), key="sel_av")
        stk = st.selectbox("📦 정상재고", cols, index=auto_idx(['정상재고']), key="sel_stk")
        
        # 3일 판매: '3일 발주합계' 최우선
        t3_target = "3일 발주합계"
        t3_idx = cols.index(t3_target) if t3_target in cols else auto_idx(['3일'], exclude_keys=['1주', '7일', '품절'])
        t3 = st.selectbox("🔥 3일 판매", cols, index=t3_idx, key="sel_t3")
        
        # 7일 판매: '1주발주합계' 최우선
        t7_target = "1주발주합계"
        t7_idx = cols.index(t7_target) if t7_target in cols else auto_idx(['7일', '1주'], exclude_keys=['3일', '품절'])
        t7 = st.selectbox("📅 7일 판매", cols, index=t7_idx, key="sel_t7")
        
        # 등록일 추가
        reg = st.selectbox("📆 상품 등록일", cols, index=auto_idx(['등록일', '등록일자', '최초등록']), key="sel_reg")

    st.write("") # 간격 조절
    
    # --- 3단계: 데이터 분석 설정 ---
    st.subheader("🚀 3단계: 데이터 분석 설정")
    
    s1, s2 = st.columns(2)
    with s1:
        lt_val = st.number_input("⏳ 리드타임 (일)", value=7, key="inp_lt")
    with s2:
        ss_val = st.number_input("🛡️ 안전재고 (일)", value=3, key="inp_ss")

    # 분석 시작 버튼
    if st.button("📊 데이터 분석 시작", use_container_width=True, type="primary"):
        st.session_state.p = {
            'so': so, 'vn': vn, 'vi': vi, 'it': it, 'op': op, 
            'st': stk, 'av': av, 't3': t3, 't7': t7, 'reg': reg,
            'lt': lt_val, 'ss': ss_val
        }
        
        df_final = st.session_state.df_raw.copy()
        
        # 등록일 날짜 형식 변환
        if reg in df_final.columns:
            df_final[reg] = pd.to_datetime(df_final[reg], errors='coerce')
        
        st.session_state.df_raw = df_final 
        st.session_state.analyzed = True   
        st.rerun()





# ------------------------------------------------------------------
# [4단계: 데이터 편집 및 재고 관리]
# ------------------------------------------------------------------
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📊 4단계: 데이터 편집 및 재고 관리")

    # 매핑 변수 설정
    p = st.session_state.p
    sold_out_col, item, option = p['so'], p['it'], p['op']
    vendor, v_item = p['vn'], p['vi']
    stock, avail, t3day, t7day = p['st'], p['av'], p['t3'], p['t7']
    reg_date_col = p.get('reg') 
    lt, ss = p['lt'], p['ss']

    # 데이터 복사 및 숫자 변환 (누락 방지)
    df_work = st.session_state.df_raw.copy()
    for col in [stock, avail, t3day, t7day]:
        if col in df_work.columns:
            df_work[col] = pd.to_numeric(df_work[col], errors='coerce').fillna(0).astype(int)
    
    if "리오더 수량" not in df_work.columns: df_work["리오더 수량"] = 0
    df_work["리오더 수량"] = pd.to_numeric(df_work["리오더 수량"], errors='coerce').fillna(0).astype(int)
    if "비고" not in df_work.columns: df_work["비고"] = ""
    df_work["리오더 입고수량"] = 0 

    # UI 레이아웃
    f_c1, f_c2, f_c3 = st.columns([1, 2, 1])
    with f_c1: filter_m = st.selectbox("🚦 필터", ["전체보기", "정상만", "품절만"], index=1, key="v4_main_filter")
    with f_c2: search_q = st.text_input("🔍 검색", placeholder="상품명 또는 옵션...", key="v4_main_search")
    with f_c3: hist_date_4 = st.date_input("🗓️ 입고 기록 날짜", datetime.now(KST).date(), key="v4_main_date")

    # 판매량 및 권장발주 계산 로직 (신상품 보정 포함)
    def calc_daily(row):
        t7, t3 = row[t7day], row[t3day]
        if reg_date_col and reg_date_col in row and pd.notnull(row[reg_date_col]):
            today = datetime.now(KST).date()
            reg_dt = row[reg_date_col].date() if hasattr(row[reg_date_col], 'date') else pd.to_datetime(row[reg_date_col]).date()
            days_diff = (today - reg_dt).days
            if 0 <= days_diff < 3:
                return int(round(t3 / (days_diff + 1))) if t3 > 0 else 0
        if t7 > 0: return int(round(t7 / 7))
        elif t3 > 0: return int(round(t3 / 3))
        return 0

    df_work['일판매'] = df_work.apply(calc_daily, axis=1)
    df_work['3일발주'] = (df_work['일판매'] * 3).astype(int)
    df_work['권장발주'] = ((df_work['일판매'] * (lt + ss)) - (df_work[avail] + df_work['리오더 수량'])).clip(lower=0).astype(int)

    # 필터링 적용
    is_so = df_work[sold_out_col].astype(str).str.contains('품절', na=False)
    df_f = df_work[~is_so] if filter_m == "정상만" else (df_work[is_so] if filter_m == "품절만" else df_work)
    if search_q:
        df_f = df_f[df_f[item].astype(str).str.contains(search_q, case=False) | df_f[option].astype(str).str.contains(search_q, case=False)]

    # 화면 출력용 이름 변경
    df_disp = df_f.rename(columns={sold_out_col: "상태", vendor: "공급쳐", v_item: "공급처상품명", item: "상품명", option: "옵션", stock: "정상", avail: "가용"})
    cols = ["상태", "공급쳐", "상품명", "옵션", "공급처상품명", "정상", "가용", "리오더 수량", "리오더 입고수량", "3일발주", "일판매", "권장발주"]

    with st.form("v4_master_form"):
        # ⭐ 에러 방지를 위해 alignment 옵션 제거, 나머지 설정은 풀버전
        edited_v4 = st.data_editor(df_disp[cols], use_container_width=True, hide_index=True, key="v4_editor",
                                column_config={
                                    "상품명": st.column_config.TextColumn(width=300),
                                    "공급처상품명": st.column_config.TextColumn(width=150),
                                    "리오더 입고수량": st.column_config.NumberColumn("리오더 입고", format="%d", min_value=0)
                                })
        
        if st.form_submit_button("💾 데이터 저장 및 입고 반영", use_container_width=True, type="primary"):
            user_edits = st.session_state["v4_editor"].get("edited_rows", {})
            if user_edits:
                m_sh, h_sh = get_sheet().worksheet("시트1"), get_sheet().worksheet("입고기록")
                for r_idx, changes in user_edits.items():
                    idx = df_disp.index[int(r_idx)]
                    if "리오더 수량" in changes: 
                        st.session_state.df_raw.at[idx, "리오더 수량"] = int(changes["리오더 수량"])
                    if "리오더 입고수량" in changes:
                        qty = int(changes["리오더 입고수량"])
                        if qty > 0:
                            old_v = int(st.session_state.df_raw.at[idx, "리오더 수량"])
                            st.session_state.df_raw.at[idx, "리오더 수량"] = max(0, old_v - qty)
                            h_sh.append_row([datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S'), str(df_disp.at[idx, "상품명"]), str(df_disp.at[idx, "옵션"]), qty])
                
                df_save = st.session_state.df_raw.copy().fillna("").astype(str)
                m_sh.update([df_save.columns.values.tolist()] + df_save.values.tolist())
                st.success("✅ 메인 데이터 저장 완료!"); time.sleep(0.5); st.rerun()





# ------------------------------------------------------------------
# [5단계: 최종 발주 요약 및 엑셀 다운로드]
# ------------------------------------------------------------------
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📋 5단계: 최종 발주 리스트 요약")

    # 4단계에서 계산된 최신 데이터 활용
    df_v5 = df_work[~df_work[sold_out_col].astype(str).str.contains('품절', na=False)].copy()
    
    # 5단계 전용 검색 및 필터
    f1, f2, f3 = st.columns([1.5, 2, 1])
    m5_f = f1.selectbox("🚦 상태 필터", ["🚨 고위험/주의", "✅ 전체정상"], key="v5_main_filter")
    s5_q = f2.text_input("🔍 상품명 검색", key="v5_main_search")
    d5_d = f3.date_input("🗓️ 기준 날짜 (5단계)", datetime.now(KST).date(), key="v5_main_date")

    if 'add_order_dict' not in st.session_state: st.session_state.add_order_dict = {}
    df_v5['추가발주수량'] = df_v5.index.map(st.session_state.add_order_dict).fillna(0).astype(int)

    # 위험 상품군(권장발주 > 0) 추출
    danger_names = df_v5[df_v5['권장발주'] > 0][item].unique()
    df_v5_view = df_v5[df_v5[item].isin(danger_names)].copy() if m5_f == "🚨 고위험/주의" else df_v5[~df_v5[item].isin(danger_names)].copy()
    
    if s5_q:
        df_v5_view = df_v5_view[df_v5_view[item].astype(str).str.contains(s5_q, case=False)]
    
    df_v5_view = df_v5_view.sort_values(by=[item, option])

    # 5단계 에디터 출력
    map_v5 = {item: "상품명", option: "옵션", vendor: "공급쳐", v_item: "공급처상품명", avail: "가용", "리오더 수량": "기존리오더", "추가발주수량": "추가발주수량", "권장발주": "권장발주", "비고": "이슈/입고메모"}
    
    with st.form("v5_master_form"):
        df_ed_v5 = df_v5_view[list(map_v5.keys())].rename(columns=map_v5)
        st.data_editor(df_ed_v5, use_container_width=True, hide_index=True, key="v5_editor",
                        column_config={
                            "상품명": st.column_config.TextColumn(width=280),
                            "공급처상품명": st.column_config.TextColumn(width=150),
                            "추가발주수량": st.column_config.NumberColumn(min_value=0)
                        })
        
        if st.form_submit_button("✅ 추가발주 및 메모 확정 (메인 반영)", use_container_width=True, type="primary"):
            edits_v5 = st.session_state["v5_editor"].get("edited_rows", {})
            if edits_v5:
                for r_idx, val in edits_v5.items():
                    idx = df_v5_view.index[int(r_idx)]
                    if "추가발주수량" in val:
                        st.session_state.df_raw.at[idx, "리오더 수량"] += int(val["추가발주수량"])
                        st.session_state.add_order_dict[idx] = int(val["추가발주수량"])
                    if "이슈/입고메모" in val:
                        st.session_state.df_raw.at[idx, "비고"] = str(val["이슈/입고메모"])
                
                m_sh = get_sheet().worksheet("시트1")
                df_s = st.session_state.df_raw.copy().fillna("").astype(str)
                m_sh.update([df_s.columns.values.tolist()] + df_s.values.tolist())
                st.success("✅ 최종 발주 내용이 메인 시트에 저장되었습니다!"); time.sleep(0.5); st.rerun()

    # 하단 버튼부 (10열 저장 및 엑셀 다운로드)
    st.divider()
    c_save, c_down = st.columns(2)
    
    with c_save:
        if st.button("💾 구글 시트에 최종 발주 기록 저장", use_container_width=True):
            ready = df_v5_view[df_v5_view.index.isin(st.session_state.add_order_dict.keys())].copy()
            if not ready.empty:
                now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
                log_rows = [[
                    now_s, str(r[item]), str(r[option]), str(r[vendor]), int(r[avail]),
                    int(st.session_state.df_raw.at[i, '리오더 수량']), 
                    int(st.session_state.add_order_dict[i]), int(r['권장발주']),
                    str(st.session_state.df_raw.at[i, '비고']), str(r[v_item]) # ⭐ 10열(공급처상품명)
                ] for i, r in ready.iterrows()]
                
                try:
                    get_sheet().worksheet("발주기록").append_rows(log_rows)
                    st.session_state.add_order_dict = {} 
                    st.success("✅ 발주기록 시트에 10열 저장 완료!"); time.sleep(0.5); st.rerun()
                except Exception as e: st.error(f"❌ 저장 실패: {e}")

    with c_down:
        # 권장발주 + 추가발주가 있는 것만 다운로드
        df_v5['최종합계'] = df_v5['권장발주'] + df_v5['추가발주수량']
        csv_target = df_v5[df_v5['최종합계'] > 0].copy()
        if not csv_target.empty:
            csv_res = csv_target[[item, option, vendor, v_item, '최종합계']].rename(columns={
                item: "상품명", option: "옵션", vendor: "공급처", v_item: "공급처상품명"
            })
            csv_data = csv_res.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button("📥 최종 발주서 CSV 다운로드", csv_data, f"발주서_{d5_d.strftime('%m%d')}.csv", "text/csv", use_container_width=True)
        else:
            st.button("📥 최종 발주서 CSV 다운로드 (내역 없음)", disabled=True, use_container_width=True)
            

# ==========================================================
# --- [6단계: 전체 히스토리 관리 (열 번호 강제 매핑 교정)] ---
# ==========================================================
if st.session_state.get('analyzed'):
    st.divider()
    st.subheader("📜 6단계: 전체 히스토리 관리")

    f1, f2, f3, f4 = st.columns([1, 0.5, 1.2, 1.2])
    with f1:
        today = datetime.now(KST).date()
        d_range = st.date_input("🗓️ 조회 날짜 범위", value=(today, today), key="v6_date_final")
    with f2:
        st.write(""); st.write("") 
        search_trigger = st.button("🔍 조회하기", use_container_width=True, type="primary", key="v6_search_btn")
    with f3:
        h_q = st.text_input("🔍 결과 내 상품명 검색", placeholder="상품명을 입력하세요...", key="v6_search_final")
    with f4:
        selected_batch = st.selectbox("📥 저장 회차 선택", ["전체보기"], key="v6_batch_select")

    if search_trigger or h_q:
        try:
            with st.spinner("📡 데이터를 불러오는 중..."):
                sheet = get_sheet()
                worksheet = sheet.worksheet("발주기록")
                all_values = worksheet.get_all_values()
            
            if len(all_values) > 1:
                # 🚨 [중요] 시트의 실제 저장 순서와 1:1 매칭 (5단계 저장 리스트 순서와 동일하게)
                raw_data = all_values[1:]
                
                # 시트에서 0번부터 9번까지 순서대로 이름을 붙여줍니다.
                temp_df = pd.DataFrame(raw_data)
                
                # 사장님 시트의 실제 데이터 배치:
                # 0:날짜시간, 1:상품명, 2:옵션, 3:업체명(공급처), 4:가용재고, 
                # 5:기존리오더, 6:추가발주수량, 7:권장발주수량, 8:비고(이슈), 9:공급처상품명
                
                df_hist = pd.DataFrame()
                df_hist["날짜시간"] = temp_df[0]
                df_hist["상품명"] = temp_df[1]
                df_hist["옵션"] = temp_df[2]
                df_hist["업체명"] = temp_df[3]
                df_hist["가용재고"] = temp_df[4]
                df_hist["기존리오더"] = temp_df[5]
                df_hist["추가발주수량"] = temp_df[6]
                df_hist["권장발주수량"] = temp_df[7]
                df_hist["이슈/메모"] = temp_df[8]
                df_hist["공급처상품명"] = temp_df[9]

                # --- [필터링] ---
                df_hist["날짜_만"] = df_hist["날짜시간"].astype(str).str.slice(0, 10)
                if len(d_range) == 2:
                    s_date, e_date = d_range[0].strftime('%Y-%m-%d'), d_range[1].strftime('%Y-%m-%d')
                    df_hist = df_hist[(df_hist["날짜_만"] >= s_date) & (df_hist["날짜_만"] <= e_date)]

                df_hist["추가발주수량"] = pd.to_numeric(df_hist["추가발주수량"], errors='coerce').fillna(0)
                df_hist = df_hist[df_hist["추가발주수량"] > 0] 
                df_hist = df_hist.sort_values(by="날짜시간", ascending=False)

                if h_q:
                    df_hist = df_hist[df_hist["상품명"].astype(str).str.contains(h_q, case=False)]

                if not df_hist.empty:
                    # 🚨 [사장님 요청 순서] 발주시간 => 업체명 => 상품명 => 옵션 => 공급처 상품명 ...
                    display_order = [
                        "날짜시간", "업체명", "상품명", "옵션", "공급처상품명", 
                        "가용재고", "기존리오더", "추가발주수량", "권장발주수량", "이슈/메모"
                    ]
                    df_view = df_hist[display_order].copy()
                    df_view["이슈/메모"] = df_view["이슈/메모"].astype(str).replace('', '-')

                    st.success(f"✅ 총 **{len(df_view)}**건의 내역이 조회되었습니다.")
                    st.dataframe(
                        df_view, 
                        use_container_width=True, 
                        hide_index=True,
                        column_config={
                            "날짜시간": st.column_config.TextColumn("📅 발주시간", width="medium"),
                            "업체명": st.column_config.TextColumn("🏭 업체명", width="small"),
                            "공급처상품명": st.column_config.TextColumn("🆔 공급처 상품명", width="medium"),
                        }
                    )
                    
                    csv_data = df_view.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
                    st.download_button("📥 결과 다운로드(CSV)", csv_data, f"발주히스토리.csv", use_container_width=True)
                else:
                    st.warning("🧐 해당 조건의 기록이 없습니다.")
        except Exception as e:
            st.error(f"📡 오류: {e}")

# ==========================================================
# --- [7단계: 미입고 현황 및 메모 실시간 저장 기능] ---
# ==========================================================
if st.session_state.get('analyzed'):
    st.divider()
    st.subheader("📊 7단계: 공장 미입고 총 수량 현황")

    try:
        sh_log = get_sheet().worksheet("발주기록")
        raw_data = sh_log.get_all_values()
        
        if len(raw_data) > 1:
            # 전체 데이터를 가져옵니다 (수정을 위해 index를 포함한 원본 형태 유지)
            df_log_all = pd.DataFrame(raw_data[1:], columns=raw_data[0])
            
            # [화면 표시용 가공]
            # F(기존), G(추가) 수량을 합산하여 '미입고 총수량' 생성
            v_exist = pd.to_numeric(df_log_all.iloc[:, 5], errors='coerce').fillna(0).astype(int)
            v_added = pd.to_numeric(df_log_all.iloc[:, 6], errors='coerce').fillna(0).astype(int)
            df_log_all["미입고 총수량"] = v_exist + v_added

            # 필터/검색용 레이아웃
            col1, col2 = st.columns([1, 1])
            with col1:
                v_list = ["전체"] + sorted([v for v in df_log_all.iloc[:, 9].unique() if v])
                selected_v = st.selectbox("🏠 업체별 필터", v_list, key="v7_save_v")
            with col2:
                search_v = st.text_input("🔍 상품명 검색", key="v7_save_s")

            # 필터링 적용
            filtered_df = df_log_all.copy()
            if selected_v != "전체":
                filtered_df = filtered_df[filtered_df.iloc[:, 9] == selected_v]
            if search_v:
                mask = filtered_df.apply(lambda row: row.astype(str).str.contains(search_v, case=False).any(), axis=1)
                filtered_df = filtered_df[mask]

            # [핵심: 데이터 에디터]
            # 여기서 수정하면 edited_rows에 변경 내용이 담깁니다.
            edited_df = st.data_editor(
                filtered_df,
                use_container_width=True,
                hide_index=False, # 수정을 위해 인덱스 잠시 표시
                key="v7_editor_with_save",
                column_order=("날짜", "업체명", "상품명", "옵션", "미입고 총수량", "메모"),
                column_config={
                    "미입고 총수량": st.column_config.NumberColumn("🔢 미입고 총수량", disabled=True), # 수량은 수정 불가
                    "메모": st.column_config.TextColumn("📝 입고/이슈 메모 (수정 후 저장 버튼 클릭)", width=500),
                }
            )

            # [저장 버튼 로직]
            if st.button("💾 변경된 메모 시트에 저장하기", use_container_width=True, type="primary"):
                with st.spinner("구글 시트에 메모를 기록 중입니다..."):
                    # 변경된 셀 정보 확인
                    if st.session_state["v7_editor_with_save"]["edited_rows"]:
                        updates = st.session_state["v7_editor_with_save"]["edited_rows"]
                        
                        for row_idx, changes in updates.items():
                            if "메모" in changes:
                                # 실제 시트의 행 번호 계산 (헤더 1줄 + 인덱스는 0부터 시작하므로 +2)
                                # 주의: 필터링된 상태이므로 원본 df의 인덱스를 찾아야 함
                                actual_row_in_sheet = int(filtered_df.index[int(row_idx)]) + 2
                                new_memo = changes["메모"]
                                
                                # I열(9번째 열)이 메모 칸입니다.
                                sh_log.update_cell(actual_row_in_sheet, 9, new_memo)
                        
                        st.success("✅ 메모가 구글 시트에 안전하게 저장되었습니다!")
                        st.rerun() # 저장 후 새로고침해서 반영
                    else:
                        st.info("수정된 메모 내용이 없습니다.")

        else:
            st.warning("기록된 데이터가 없습니다.")
            
    except Exception as e:
        st.error(f"메모 저장 중 오류 발생: {e}")
