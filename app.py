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
# [5단계: 최종 발주 요약 및 구글 시트 저장]
# ------------------------------------------------------------------
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📋 5단계: 최종 발주 리스트 요약")

    # 품절 제외 데이터 구성
    df_v5 = df_work[~df_work[sold_out_col].astype(str).str.contains('품절', na=False)].copy()
    
    # 상단 필터
    f1, f2, f3 = st.columns([1.5, 2, 1])
    m5_f = f1.selectbox("🚦 상태 필터", ["🚨 고위험/주의", "✅ 전체정상"], key="v5_main_filter")
    s5_q = f2.text_input("🔍 상품명 검색", key="v5_main_search", placeholder="상품명 입력...")
    d5_d = f3.date_input("🗓️ 기준 날짜", datetime.now(KST).date(), key="v5_main_date")

    if 'add_order_dict' not in st.session_state: 
        st.session_state.add_order_dict = {}
    
    df_v5['추가발주수량'] = df_v5.index.map(st.session_state.add_order_dict).fillna(0).astype(int)
    df_v5['상태'] = df_v5['권장발주'].apply(lambda x: "🚨 긴급" if x > 0 else "✅ 정상")

    # 필터링 로직
    danger_names = df_v5[df_v5['권장발주'] > 0][item].unique()
    if m5_f == "🚨 고위험/주의":
        df_v5_view = df_v5[df_v5[item].isin(danger_names)].copy()
    else:
        df_v5_view = df_v5[~df_v5[item].isin(danger_names)].copy()
    
    if s5_q:
        df_v5_view = df_v5_view[df_v5_view[item].astype(str).str.contains(s5_q, case=False)]
    
    df_v5_view = df_v5_view.sort_values(by=[item, option])

    # 에디터 출력
    map_v5 = {
        "상태": "상태", item: "상품명", option: "옵션", v_item: "공급처상품명", 
        avail: "가용", "리오더 수량": "기존", "추가발주수량": "추가", "권장": "권장", "비고": "📝 이슈/입고메모"
    }
    
    with st.form("v5_master_form"):
        df_ed_v5 = df_v5_view[list(map_v5.keys())].rename(columns=map_v5)
        st.data_editor(
            df_ed_v5, use_container_width=True, hide_index=True, key="v5_editor",
            column_config={
                "상태": st.column_config.TextColumn(width=70),
                "상품명": st.column_config.TextColumn(width=250),
                "옵션": st.column_config.TextColumn(width=150),
                "공급처상품명": st.column_config.TextColumn(width=180),
                "가용": st.column_config.NumberColumn(width=80),
                "기존": st.column_config.NumberColumn(width=80),
                "추가": st.column_config.NumberColumn(width=80, min_value=0),
                "권장": st.column_config.NumberColumn(width=80),
                "📝 이슈/입고메모": st.column_config.TextColumn(width=450)
            },
            disabled=["상태", "상품명", "옵션", "공급처상품명", "가용", "기존", "권장"]
        )
        
        if st.form_submit_button("✅ 1. 추가발주 및 메모 확정 (메인 반영)", use_container_width=True, type="primary"):
            edits = st.session_state["v5_editor"].get("edited_rows", {})
            if edits:
                for r_idx, val in edits.items():
                    idx = df_v5_view.index[int(r_idx)]
                    if "추가" in val:
                        st.session_state.df_raw.at[idx, "리오더 수량"] += int(val["추가"])
                        st.session_state.add_order_dict[idx] = int(val["추가"])
                    if "📝 이슈/입고메모" in val:
                        st.session_state.df_raw.at[idx, "비고"] = str(val["📝 이슈/입고메모"])
                
                # 메인 시트1 업데이트
                m_sh = get_sheet().worksheet("시트1")
                df_s = st.session_state.df_raw.copy().fillna("").astype(str)
                m_sh.update([df_s.columns.values.tolist()] + df_s.values.tolist())
                st.success("✅ 메인 시트 반영 완료!"); time.sleep(0.5); st.rerun()

    # 하단 저장 및 다운로드
    st.divider()
    c_save, c_down = st.columns(2)
    with c_save:
        if st.button("💾 2. 구글 시트에 최종 발주 기록 저장", use_container_width=True, type="primary"):
            # 확정된 데이터가 있는지 체크
            valid_dict = {k: v for k, v in st.session_state.add_order_dict.items() if v > 0}
            if not valid_dict:
                st.warning("⚠️ 저장할 추가 수량이 없습니다. [확정] 버튼을 먼저 눌러주세요.")
            else:
                try:
                    with st.spinner("🚀 발주기록 시트로 전송 중..."):
                        ws_log = get_sheet().worksheet("발주기록")
                        now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
                        log_rows = []
                        for i, qty in valid_dict.items():
                            row = st.session_state.df_raw.loc[i]
                            log_rows.append([
                                now_s, str(row[item]), str(row[option]), str(row[v_item]), # 3번:거래처상품명
                                int(row[avail]), int(row['리오더 수량']-qty), int(qty), 
                                int(row['권장발주']), str(row['비고']), str(row[vendor]) # 9번:업체명
                            ])
                        ws_log.append_rows(log_rows)
                        st.session_state.add_order_dict = {} 
                        st.success("✅ 발주기록 저장 완료!"); time.sleep(1); st.rerun()
                except Exception as e: st.error(f"❌ 저장 실패: {e}")

    with c_down:
        df_v5['최종합계'] = df_v5['권장발주'] + df_v5['추가발주수량']
        csv_target = df_v5[df_v5['최종합계'] > 0].copy()
        if not csv_target.empty:
            csv_res = csv_target[[vendor, item, option, v_item, '최종합계']].rename(columns={
                vendor: "공급처", item: "상품명", option: "옵션", v_item: "공급처상품명", "최종합계": "발주수량"
            })
            csv_data = csv_res.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button("📥 최종 발주서 CSV 다운로드", csv_data, f"발주서_{d5_d.strftime('%m%d')}.csv", use_container_width=True)
