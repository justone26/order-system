import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- [1. 함수 정의 섹션] ---
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

def get_auto_index(cols, keywords):
    for key in keywords:
        for i, c in enumerate(cols):
            if key in str(c): return i
    return 0

# --- [2. 앱 설정 및 초기화] ---
st.set_page_config(layout="wide", page_title="재고 관리 시스템")
st.title("📦 재고 관리 및 발주 시스템")

if 'history' not in st.session_state: st.session_state.history = {}

# --- [3. 데이터 업로드 및 초기화] ---
st.subheader("📁 데이터 업로드")
uploaded_file = st.file_uploader("엑셀/CSV 파일을 여기에 드래그하거나 선택하세요", type=['xlsx', 'xls', 'csv'])

if st.button("🔄 전체 데이터 초기화 및 재업로드"):
    st.session_state.clear()
    st.rerun()

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
    
    # c1 컬럼에 주요 정보 배치
    sold_out = c1.selectbox("품절 여부", cols, index=get_auto_index(cols, ['품절']))
    vendor = c1.selectbox("공급처", cols, index=get_auto_index(cols, ['공급처']))
    item = c1.selectbox("상품명", cols, index=get_auto_index(cols, ['상품명']))
    option = c1.selectbox("옵션", cols, index=get_auto_index(cols, ['옵션']))
    vendor_item = c1.selectbox("공급처 상품명", cols, index=get_auto_index(cols, ['공급처상품명']))
    
    # c2 컬럼에 수치 및 날짜 정보 배치
    reg_date = c2.selectbox("등록일", cols, index=get_auto_index(cols, ['등록일'])) # 여기에 등록일 추가!
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
        df['권장 발주량'] = ((daily_avg * lead_time) + (daily_avg * safety_stock) - pd.to_numeric(df[avail], errors='coerce')).clip(lower=0).astype(int)
        st.session_state.df_raw = df
        st.rerun()

    # 4단계: 데이터 편집
    st.subheader("📊 4단계: 데이터 편집")
    df_working = st.session_state.df_raw.copy()
    st.data_editor(df_working, use_container_width=True, key="main_editor")

    # 5단계: 발주 리스트 요약
    st.subheader("📋 5단계: 발주 리스트 요약")
    if '권장 발주량' in st.session_state.df_raw.columns:
        to_order = st.session_state.df_raw[st.session_state.df_raw['권장 발주량'] > 0].copy()
        if not to_order.empty:
            to_order['추가 리오더'] = 0
            display_cols = [vendor, item, option, "리오더 수량", "추가 리오더", "권장 발주량"]
            edited = st.data_editor(to_order[display_cols], use_container_width=True, key="order_editor")
            
            def highlight_urgent(row):
                avail_val = float(row.get('가용재고', 0))
                sale_3d = float(row.get(t3day, 0)) / 3
                return ['background-color: #ffcccc'] * len(row) if avail_val <= sale_3d else [''] * len(row)
            
            st.dataframe(edited.style.apply(highlight_urgent, axis=1), use_container_width=True)
            
            col1, col2 = st.columns(2)
            if col1.button("💾 구글 시트 및 기록 저장"):
                for idx, row in edited.iterrows():
                    st.session_state.df_raw.at[idx, '리오더 수량'] = float(row['리오더 수량']) + float(row['추가 리오더'])
                save_reorder_data(st.session_state.df_raw.loc[edited.index, ['상품코드', '리오더 수량']])
                st.session_state.history[datetime.now().strftime("%Y-%m-%d %H:%M:%S")] = edited.copy()
                st.success("✅ 저장 완료!")
            
            with col2:
                csv_data = edited.to_csv(index=False).encode('utf-8-sig')
                st.download_button("📥 발주 리스트 엑셀 다운로드", csv_data, f"발주_{datetime.now().strftime('%Y%m%d')}.csv", "text/csv")

    # 6단계: 과거 데이터 확인
    st.subheader("📜 6단계: 과거 데이터 확인")
    if st.session_state.history:
        select_h = st.selectbox("⏰ 기록된 시간 선택", list(st.session_state.history.keys()))
        st.dataframe(st.session_state.history[select_h], use_container_width=True)
        csv_h = st.session_state.history[select_h].to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 선택 기록 다운로드", csv_h, f"발주기록_{select_h.replace(':', '-')}.csv", "text/csv")
    else:
        st.info("💡 아직 저장된 기록이 없습니다.")

