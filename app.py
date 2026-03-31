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
# [5단계: 최종 발주 요약 - KeyError 방지 패치 완료]
# ------------------------------------------------------------------
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📋 5단계: 최종 발주 리스트 요약")

    # [수정] df_work가 정의되지 않았을 경우를 대비해 세션 데이터에서 직접 복사
    df_v5 = st.session_state.df_raw.copy()
    
    # 품절 제외 필터링 (sold_out_col 변수 체크)
    if 'sold_out_col' in globals() and sold_out_col in df_v5.columns:
        df_v5 = df_v5[~df_v5[sold_out_col].astype(str).str.contains('품절', na=False)]
    
    # 상단 필터 레이아웃
    f1, f2, f3 = st.columns([1.5, 2, 1])
    m5_f = f1.selectbox("🚦 상태 필터", ["🚨 고위험/주의", "✅ 전체정상"], key="v5_main_filter")
    s5_q = f2.text_input("🔍 상품명 검색", key="v5_main_search", placeholder="상품명 입력...")
    d5_d = f3.date_input("🗓️ 기준 날짜", datetime.now(KST).date(), key="v5_main_date")

    if 'add_order_dict' not in st.session_state: 
        st.session_state.add_order_dict = {}
    
    # 컬럼 존재 확인 후 계산
    df_v5['추가발주수량'] = df_v5.index.map(st.session_state.add_order_dict).fillna(0).astype(int)
    
    # 권장발주 컬럼이 없을 경우 0으로 생성 (KeyError 방지)
    if '권장발주' not in df_v5.columns: df_v5['권장발주'] = 0
    df_v5['상태'] = df_v5['권장발주'].apply(lambda x: "🚨 긴급" if x > 0 else "✅ 정상")

    # 필터링 로직
    danger_names = df_v5[df_v5['권장발주'] > 0][item].unique() if item in df_v5.columns else []
    if m5_f == "🚨 고위험/주의":
        df_v5_view = df_v5[df_v5[item].isin(danger_names)].copy() if item in df_v5.columns else df_v5.copy()
    else:
        df_v5_view = df_v5[~df_v5[item].isin(danger_names)].copy() if item in df_v5.columns else df_v5.copy()
    
    if s5_q and item in df_v5_view.columns:
        df_v5_view = df_v5_view[df_v5_view[item].astype(str).str.contains(s5_q, case=False)]
    
    # 정렬
    sort_cols = [c for c in [item, option] if c in df_v5_view.columns]
    if sort_cols: df_v5_view = df_v5_view.sort_values(by=sort_cols)

    # 🚨 [중요] KeyError 방지용 매핑 체크
    # map_v5의 키값들이 실제 df_v5_view에 존재하는지 확인된 것만 추출합니다.
    full_map = {
        "상태": "상태", item: "상품명", option: "옵션", v_item: "공급처상품명", 
        avail: "가용", "리오더 수량": "기존", "추가발주수량": "추가", "권장발주": "권장", "비고": "📝 이슈/입고메모"
    }
    # 실제 존재하는 컬럼만 골라내기 (이게 없으면 KeyError 발생함)
    available_map = {k: v for k, v in full_map.items() if k in df_v5_view.columns}
    
    with st.form("v5_master_form"):
        # 존재하는 컬럼으로만 에디터 생성
        df_ed_v5 = df_v5_view[list(available_map.keys())].rename(columns=available_map)
        
        st.data_editor(
            df_ed_v5, use_container_width=True, hide_index=True, key="v5_editor",
            column_config={
                "상품명": st.column_config.TextColumn(width=250),
                "옵션": st.column_config.TextColumn(width=150),
                "추가": st.column_config.NumberColumn(width=80, min_value=0),
                "📝 이슈/입고메모": st.column_config.TextColumn(width=450)
            },
            disabled=[col for col in df_ed_v5.columns if col != "추가" and col != "📝 이슈/입고메모"]
        )
        
        if st.form_submit_button("✅ 추가발주 및 메모 확정 (메인 반영)", use_container_width=True, type="primary"):
            edits = st.session_state["v5_editor"].get("edited_rows", {})
            if edits:
                for r_idx, val in edits.items():
                    idx = df_v5_view.index[int(r_idx)]
                    if "추가" in val:
                        st.session_state.df_raw.at[idx, "리오더 수량"] = int(st.session_state.df_raw.at[idx, "리오더 수량"]) + int(val["추가"])
                        st.session_state.add_order_dict[idx] = int(val["추가"])
                    if "📝 이슈/입고메모" in val:
                        st.session_state.df_raw.at[idx, "비고"] = str(val["📝 이슈/입고메모"])
                
                m_sh = get_sheet().worksheet("시트1")
                df_s = st.session_state.df_raw.copy().fillna("").astype(str)
                m_sh.update([df_s.columns.values.tolist()] + df_s.values.tolist())
                st.success("✅ 메인 시트 반영 완료!"); time.sleep(0.5); st.rerun()




# ------------------------------------------------------------------
# [6단계: 전체 히스토리 관리 - 상단 필터 상시 노출 버전]
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.subheader("📜 6단계: 전체 히스토리 관리")

    f1, f2, f3, f4 = st.columns([1.2, 0.8, 1.5, 1.5])
    with f1: d_range_v6 = st.date_input("🗓️ 1. 조회 범위", value=(today, today), key="v6_range")
    with f2: 
        st.write(""); st.write("")
        btn_v6 = st.button("🔍 2. 내역 조회", use_container_width=True, type="primary")

    # 세션 데이터 로드
    if btn_v6:
        all_h = get_sheet().worksheet("발주기록").get_all_values()
        if len(all_h) > 1:
            df_v6_raw = pd.DataFrame(all_h[1:])
            df_v6_raw.columns = ["발주시간", "상품명", "옵션", "공급처상품명", "가용", "기존", "추가", "권장", "이슈/메모", "업체명"]
            df_v6_raw["날짜"] = df_v6_raw["발주시간"].str.slice(0, 10)
            s_d, e_d = d_range_v6[0].strftime('%Y-%m-%d'), d_range_v6[1].strftime('%Y-%m-%d')
            st.session_state.v6_data = df_v6_raw[(df_v6_raw["날짜"] >= s_d) & (df_v6_raw["날짜"] <= e_d)]
            st.session_state.v6_times = sorted(st.session_state.v6_data["발주시간"].unique(), reverse=True)

    with f3: v6_q = st.text_input("🔍 3. 상품명 검색", key="v6_q")
    with f4:
        v6_times = st.session_state.get('v6_sessions', [])
        opts = ["📊 일자별 전체 합계"] + [f"{len(v6_times)-i}회차 ({t[11:16]})" for i, t in enumerate(v6_times)] if v6_times else ["조회 결과 없음"]
        sel_v6 = st.selectbox("📦 4. 회차 선택", opts, key="v6_sel")

    if st.session_state.get('v6_data') is not None:
        df_v6 = st.session_state.v6_data.copy()
        if sel_v6 == "📊 일자별 전체 합계":
            df_v6_disp = df_v6.groupby(["업체명", "상품명", "옵션", "공급처상품명"], as_index=False).agg({"발주시간":"max","추가":lambda x: pd.to_numeric(x).sum(),"이슈/메모":" / ".join, "가용":"last","기존":"last","권장":"last"})
        else:
            t_idx = opts.index(sel_v6) - 1
            df_v6_disp = df_v6[df_v6["발주시간"] == v6_times[t_idx]]
        
        if v6_q: df_v6_disp = df_v6_disp[df_v6_disp["상품명"].str.contains(v6_q, case=False)]
        
        st.dataframe(df_v6_disp[["발주시간", "업체명", "상품명", "옵션", "공급처상품명", "가용", "기존", "추가", "권장", "이슈/메모"]], 
                     use_container_width=True, hide_index=True,
                     column_config={"발주시간": st.column_config.TextColumn(width=160), "추가": st.column_config.NumberColumn(width=80), "이슈/메모": st.column_config.TextColumn(width=450)})


# ------------------------------------------------------------------
# [7단계: 실시간 리오더 누적 상황판 - 상단 보드 + 업체선택 포함]
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.subheader("🚀 7단계: 실시간 리오더 누적 및 상황판")

    try:
        ws_v7 = get_sheet().worksheet("발주기록")
        df_v7_raw = pd.DataFrame(ws_v7.get_all_values()[1:])
        df_v7_raw.columns = ["발주시간", "상품명", "옵션", "거래처상품명", "가용", "기존", "추가", "권장", "이슈/메모", "업체명"]
        df_v7_raw["날짜"] = df_v7_raw["발주시간"].str.slice(0, 10)
        df_v7_raw["추가"] = pd.to_numeric(df_v7_raw["추가"], errors='coerce').fillna(0)

        # 상단 상황판 (오늘 기준)
        df_today = df_v7_raw[df_v7_raw["날짜"] == datetime.now(KST).strftime('%Y-%m-%d')]
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("📅 기준 날짜", datetime.now(KST).strftime('%Y-%m-%d'))
        c2.metric("📦 상품수", f"{df_today['상품명'].nunique()}종")
        c3.metric("🏭 업체수", f"{df_today['업체명'].nunique()}곳")
        c4.metric("🔢 누적수량", f"{int(df_today['추가'].sum())}개")
        st.divider()

        # 필터 레이아웃 (기간 -> 갱신 -> 검색 -> 업체)
        f1, f2, f3, f4 = st.columns([1.2, 0.8, 1.5, 1.5])
        d_range_v7 = f1.date_input("🗓️ 1. 기간 선택", value=(today, today), key="v7_range")
        f2.write(""); f2.write(""); btn_v7 = f2.button("📈 2. 데이터 갱신", use_container_width=True, type="primary")
        q_v7 = f3.text_input("🔍 3. 상품명 검색", key="v7_q")
        v_list = sorted(df_v7_raw["업체명"].unique().tolist())
        v_choice = f4.selectbox("🏭 4. 업체 선택", ["전체 업체"] + v_list, key="v7_v_sel")

        # 필터링 및 집계
        s_d, e_d = d_range_v7[0].strftime('%Y-%m-%d'), d_range_v7[1].strftime('%Y-%m-%d')
        df_f = df_v7_raw[(df_v7_raw["날짜"] >= s_d) & (df_v7_raw["날짜"] <= e_d)]
        if q_v7: df_f = df_f[df_f["상품명"].str.contains(q_v7, case=False)]
        if v_choice != "전체 업체": df_f = df_f[df_f["업체명"] == v_choice]

        df_res = df_f.groupby(["날짜", "업체명", "상품명", "옵션", "거래처상품명"], as_index=False).agg({"추가":"sum", "이슈/메모":" / ".join})
        
        # 출력: 날짜 -> 업체명 -> 상품명 -> 옵션 -> 거래처상품명 -> 추가
        st.dataframe(df_res[["날짜", "업체명", "상품명", "옵션", "거래처상품명", "추가", "이슈/메모"]], 
                     use_container_width=True, hide_index=True,
                     column_config={"추가": st.column_config.NumberColumn("🔢 누적수량", width=80), "이슈/메모": st.column_config.TextColumn(width=400)})
    except Exception as e: st.error(f"📡 오류: {e}")



