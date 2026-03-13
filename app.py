import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- 함수 정의 ---
def get_sheet():
    creds_dict = dict(st.secrets["gcp_service_account"])
    scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    spreadsheet_key = "1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg"
    return client.open_by_key(spreadsheet_key).sheet1

def load_reorder_data():
    try:
        sheet = get_sheet()
        return pd.DataFrame(sheet.get_all_records())
    except: return pd.DataFrame(columns=['상품코드', '1차 리오더', '2차 리오더'])

def get_auto_index(cols, keywords):
    for key in keywords:
        for i, c in enumerate(cols):
            if key in str(c): return i
    return 0

# --- 앱 실행 ---
st.set_page_config(layout="wide", page_title="재고 관리 시스템")
st.title("📦 재고 관리 및 발주 시스템")

# [파일 업로드 및 초기화 로직]
uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'xls', 'csv'])

if uploaded_file is not None:
    if 'last_filename' not in st.session_state or st.session_state.last_filename != uploaded_file.name:
        df = pd.read_excel(uploaded_file)
        st.session_state.df_raw = df.loc[:, ~df.columns.duplicated()]
        st.session_state.last_filename = uploaded_file.name
        st.rerun()

if st.session_state.get('df_raw') is not None:
    cols = st.session_state.df_raw.columns.tolist()

    # 1단계: 매핑 설정
    st.subheader("⚙️ 1단계: 자동 매핑 설정")
    c1, c2 = st.columns(2)
    sold_out = c1.selectbox("품절 여부", cols, index=get_auto_index(cols, ['품절']))
    vendor = c1.selectbox("공급처", cols, index=get_auto_index(cols, ['공급처']))
    item = c1.selectbox("상품명", cols, index=get_auto_index(cols, ['상품명']))
    option = c1.selectbox("옵션", cols, index=get_auto_index(cols, ['옵션']))
    vendor_item = c1.selectbox("공급처 상품명", cols, index=get_auto_index(cols, ['공급처상품명']))
    
    reg_date = c2.selectbox("등록일", cols, index=get_auto_index(cols, ['등록일']))
    stock = c2.selectbox("정상재고", cols, index=get_auto_index(cols, ['정상재고']))
    avail = c2.selectbox("가용재고", cols, index=get_auto_index(cols, ['가용재고']))
    t3day = c2.selectbox("3일 발주합계", cols, index=get_auto_index(cols, ['3일']))
    t1week = c2.selectbox("7일 발주합계", cols, index=get_auto_index(cols, ['7일']))

    # 2~3단계: 분석
    st.subheader("⚙️ 2~3단계: 분석 설정")
    lead_time = st.number_input("리드타임 (일)", value=10)
    safety_stock = st.number_input("안전재고 (일 수)", value=7)
    if st.button("🚀 분석 실행"):
        df = st.session_state.df_raw.copy()
        daily_avg = pd.to_numeric(df[t1week], errors='coerce').fillna(0) / 7
        df['권장 발주량'] = ((daily_avg * lead_time) + (daily_avg * safety_stock) - pd.to_numeric(df[avail], errors='coerce').fillna(0)).clip(lower=0).astype(int)
        st.session_state.df_raw = df
        st.rerun()

# 4단계: 데이터 편집
    st.subheader("📊 4단계: 데이터 편집")

    target_cols = [sold_out, vendor, item, option, vendor_item, 
                   "정상재고", "가용재고", "리오더 수량", "리오더입고수량", 
                   "일판매량", "3일발주합계", "1주발주합계", "권장 발주량"]

    if "3일발주합계" in st.session_state.df_raw.columns:
        st.session_state.df_raw["일판매량"] = (pd.to_numeric(st.session_state.df_raw["3일발주합계"], errors='coerce').fillna(0) / 3).round(1)

    for c in target_cols:
        if c not in st.session_state.df_raw.columns: st.session_state.df_raw[c] = 0

    f1, f2 = st.columns([3, 1])
    search_query = f1.text_input("🔍 상품명 검색")
    filter_mode = f2.selectbox("품절 필터", ["전체보기", "정상만", "품절만"], index=1)

    df_working = st.session_state.df_raw.copy()
    if filter_mode == "정상만":
        df_working = df_working[~df_working[sold_out].astype(str).str.contains('품절', na=False)]
    elif filter_mode == "품절만":
        df_working = df_working[df_working[sold_out].astype(str).str.contains('품절', na=False)]
    if search_query:
        df_working = df_working[df_working[item].astype(str).str.contains(search_query, case=False, na=False)]

    def update_reorder():
        edited = st.session_state["main_editor"]
        for row_idx, changes in edited['edited_rows'].items():
            if '리오더입고수량' in changes:
                received = float(changes['리오더입고수량'])
                original_idx = df_working.index[row_idx]
                st.session_state.df_raw.at[original_idx, '리오더 수량'] = max(0, float(st.session_state.df_raw.at[original_idx, '리오더 수량']) - received)
                st.session_state.df_raw.at[original_idx, '리오더입고수량'] = 0

    st.data_editor(df_working[target_cols], use_container_width=True, key="main_editor", on_change=update_reorder)
                
# 5단계: 발주 리스트 요약 (수정된 코드)
    st.subheader("📋 5단계: 발주 리스트 요약")

    if '권장 발주량' in st.session_state.df_raw.columns:
        to_order = st.session_state.df_raw[st.session_state.df_raw['권장 발주량'] > 0].copy()
        
        if not to_order.empty:
            st.warning(f"🚨 발주 대상 {len(to_order)}개 확인")
            
            to_order['추가 리오더'] = 0
            cols_to_show = [vendor, item, option, "리오더 수량", "추가 리오더", "권장 발주량", "가용재고"]
            
            # 1. 스타일링을 뺀 편집기 (에러 방지)
            edited_to_order = st.data_editor(
                to_order[cols_to_show], 
                use_container_width=True, 
                key="order_editor"
            )
            
            # 2. 결과 확인 및 경고 표시 (스타일링은 여기서 적용)
            def highlight_urgent(row):
                # 3일 발주합계가 없는 경우를 대비해 0으로 처리
                val3 = float(row['3일발주합계']) if '3일발주합계' in row else 0
                is_urgent = float(row['가용재고']) <= (val3 / 3)
                return ['background-color: #ffcccc' if is_urgent else '' for _ in row]

            st.write("### 🔍 긴급 발주 확인 (가용재고 1일치 미만)")
            st.dataframe(
                edited_to_order.style.apply(highlight_urgent, axis=1), 
                use_container_width=True
            )
            
if st.button("💾 구글 시트 및 기록 저장"):
                # 1. 데이터 업데이트
                for idx, row in edited_to_order.iterrows():
                    # st.session_state.df_raw가 해당 idx를 가지고 있는지 확인
                    if idx in st.session_state.df_raw.index:
                        new_val = float(row['리오더 수량']) + float(row['추가 리오더'])
                        st.session_state.df_raw.at[idx, '리오더 수량'] = new_val
                
                # 2. 구글 시트 저장용 데이터 생성
                final_save = st.session_state.df_raw.loc[to_order.index, ['상품코드', '리오더 수량']].copy()
                final_save.columns = ['상품코드', '1차 리오더', '2차 리오더'] # 기존 함수 규격에 맞게 컬럼명 재조정
                
                # 3. 함수 호출 (함수가 코드 최상단에 있는지 확인!)
                try:
                    save_reorder_data(final_save)
                    st.success("✅ 합산 완료 및 구글 시트 저장 완료!")
                except Exception as e:
                    st.error(f"저장 중 오류 발생: {e}")
            
    # [6단계: 과거 기록]
    st.subheader("📜 6단계: 과거 데이터 확인")
    if st.session_state.history:
        select_h = st.selectbox("⏰ 시간 선택", list(st.session_state.history.keys()))
        st.dataframe(st.session_state.history[select_h])




































