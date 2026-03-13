import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- [1. 함수 정의] ---
def get_sheet():
    creds_dict = dict(st.secrets["gcp_service_account"])
    scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    spreadsheet_key = "1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg"
    return client.open_by_key(spreadsheet_key).sheet1

def save_reorder_data(df):
    try:
        sheet = get_sheet()
        sheet.clear()
        sheet.update([df.columns.values.tolist()] + df.values.tolist())
    except Exception as e:
        st.error(f"구글 시트 저장 실패: {e}")

def find_idx(cols, target_keywords):
    for keyword in target_keywords:
        for i, col in enumerate(cols):
            if keyword in str(col): return i
    return 0

# --- [2. 앱 설정 및 세션 초기화] ---
st.set_page_config(layout="wide", page_title="재고 관리 시스템")
st.title("📦 재고 관리 및 발주 시스템")
if 'history' not in st.session_state: st.session_state.history = {}

# --- [3. 데이터 업로드 및 초기화] ---
st.subheader("📁 데이터 업로드")
if st.button("🔄 전체 데이터 초기화 및 재업로드"):
    st.session_state.clear()
    st.rerun()

uploaded_file = st.file_uploader("엑셀/CSV 파일을 여기에 드래그하거나 선택하세요", type=['xlsx', 'xls', 'csv'])
st.divider()

if uploaded_file is not None:
    if 'last_filename' not in st.session_state or st.session_state.last_filename != uploaded_file.name:
        st.session_state.df_raw = pd.read_excel(uploaded_file).loc[:, ~pd.read_excel(uploaded_file).columns.duplicated()]
        st.session_state.last_filename = uploaded_file.name
        st.rerun()

# --- [4. 메인 로직] ---
if st.session_state.get('df_raw') is not None:
    cols = st.session_state.df_raw.columns.tolist()

    # 1단계: 매핑 설정
    st.subheader("⚙️ 1단계: 매핑 설정")
    c1, c2 = st.columns(2)
    sold_out = c1.selectbox("품절 여부", cols, index=find_idx(cols, ['품절']))
    vendor = c1.selectbox("공급처", cols, index=find_idx(cols, ['공급처']))
    item = c1.selectbox("상품명", cols, index=find_idx(cols, ['상품명']))
    option = c1.selectbox("옵션", cols, index=find_idx(cols, ['옵션']))
    vendor_item = c1.selectbox("공급처 상품명", cols, index=find_idx(cols, ['공급처상품명']))
    reg_date = c2.selectbox("등록일", cols, index=find_idx(cols, ['등록일']))
    stock = c2.selectbox("정상재고", cols, index=find_idx(cols, ['정상재고']))
    avail = c2.selectbox("가용재고", cols, index=find_idx(cols, ['가용재고']))
    t3day = c2.selectbox("3일 발주합계", cols, index=find_idx(cols, ['3일']))
    t1week = c2.selectbox("7일 발주합계", cols, index=find_idx(cols, ['7일', '1주']))

    # 2~3단계: 분석 설정
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
    
# 5단계: 발주 리스트 요약
    st.subheader("📋 5단계: 발주 리스트 요약")
    
    if '권장 발주량' in st.session_state.df_raw.columns:
        to_order = st.session_state.df_raw[st.session_state.df_raw['권장 발주량'] > 0].copy()
        
        if not to_order.empty:
            to_order['추가 리오더'] = 0
            display_cols = [vendor, item, option, "리오더 수량", "추가 리오더", "권장 발주량"]
            
            # 스타일링 함수: 가용재고가 1일치 판매량보다 적으면 빨간색 행 반환
            def highlight_urgent(row):
                avail_val = float(row.get('가용재고', 0))
                sale_3d = float(row.get(t3day, 0)) / 3
                return ['background-color: #ffcccc'] * len(row) if avail_val <= sale_3d else [''] * len(row)

            # 데이터 편집기를 스타일링된 DataFrame으로 감싸기
            # Note: data_editor 자체에는 직접 스타일을 입힐 수 없으므로,
            # 편집기와 스타일링된 확인용 테이블을 하나의 흐름으로 정리
            st.write("### ✏️ 발주 수량 입력 및 확인")
            
            # 편집기
            edited = st.data_editor(
                to_order[display_cols].style.apply(highlight_urgent, axis=1), 
                use_container_width=True, 
                key="order_editor"
            )
            
            # (이후 저장 및 다운로드 버튼 로직은 그대로 유지)
            col1, col2 = st.columns(2)
            if col1.button("💾 구글 시트 및 기록 저장"):
                # ... (기존 저장 로직 동일)
                st.success("✅ 저장 완료!")
            with col2:
                # ... (기존 다운로드 로직 동일)
                
    # 6단계: 과거 데이터 확인
    st.subheader("📜 6단계: 과거 데이터 확인")
    if st.session_state.history:
        select_h = st.selectbox("⏰ 기록된 시간 선택", list(st.session_state.history.keys()))
        st.dataframe(st.session_state.history[select_h], use_container_width=True)
        csv_h = st.session_state.history[select_h].to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 선택 기록 다운로드", csv_h, f"발주기록_{select_h.replace(':', '-')}.csv", "text/csv")




