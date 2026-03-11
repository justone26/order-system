import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(layout="wide", page_title="재고 관리 시스템")
st.title("📦 재고 관리 및 발주 시스템")

if 'df_raw' not in st.session_state: 
    st.session_state.df_raw = None
if 'history' not in st.session_state: 
    st.session_state.history = {}

def get_idx(cols, keywords):
    for key in keywords:
        for i, c in enumerate(cols):
            if key in str(c): 
                return i
    return 0

uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'xls', 'csv'])
if uploaded_file is not None and st.session_state.df_raw is None:
    df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
    st.session_state.df_raw = df.loc[:, ~df.columns.duplicated()]
    if "입고예정수량(리오더)" not in st.session_state.df_raw.columns:
        st.session_state.df_raw["입고예정수량(리오더)"] = 0
    st.rerun()

if st.session_state.df_raw is not None:
    cols = st.session_state.df_raw.columns.tolist()
    st.subheader("⚙️ 1단계: 자동 매핑 설정")
    c1, c2 = st.columns(2)
    with c1:
        sold_out = st.selectbox("품절 여부", cols, index=get_idx(cols, ['품절']))
        vendor = st.selectbox("공급처", cols, index=get_idx(cols, ['공급처']))
        item = st.selectbox("상품명", cols, index=get_idx(cols, ['상품명']))
    with c2:
        avail = st.selectbox("가용재고", cols, index=get_idx(cols, ['가용재고']))
        t3day = st.selectbox("3일 발주 합계", cols, index=get_idx(cols, ['3일']))

    st.subheader("⚙️ 2~3단계: 분석 실행")
    lead_time = st.number_input("리드타임", value=0)
    safety_stock = st.number_input("안전재고", value=3)
    if st.button("🚀 분석 실행"):
        st.session_state.df_raw['일일 판매량'] = (pd.to_numeric(st.session_state.df_raw[t3day], errors='coerce') / 3).round(0)
        st.session_state.df_raw['권장 발주량'] = (st.session_state.df_raw['일일 판매량'] * (lead_time + safety_stock) - (pd.to_numeric(st.session_state.df_raw[avail], errors='coerce') + st.session_state.df_raw["입고예정수량(리오더)"])).clip(lower=0)
        st.rerun()

    st.subheader("📊 4단계: 검색 및 데이터 편집")
    search = st.text_input("🔍 상품명 검색")
    df_disp = st.session_state.df_raw.copy()
    if search: 
        df_disp = df_disp[df_disp[item].astype(str).str.contains(search, na=False)]
    edited_df = st.data_editor(df_disp, use_container_width=True)
    st.session_state.df_raw.update(edited_df)

    st.subheader("📋 5단계: 발주 필요 리스트 요약")
    if '권장 발주량' in st.session_state.df_raw.columns:
        to_order = st.session_state.df_raw[st.session_state.df_raw['권장 발주량'] > 0]
        st.dataframe(to_order[[vendor, item, avail, '권장 발주량']], use_container_width=True)
        if st.button("💾 리스트 기록 저장"):
            date_key = datetime.now().strftime("%Y-%m-%d")
            record = to_order.copy()
            record['저장시각'] = datetime.now().strftime("%H:%M:%S")
            if date_key not in st.session_state.history: 
                st.session_state.history[date_key] = []
            st.session_state.history[date_key].append(record)
            st.success("저장 완료!")

    st.subheader("📜 6단계: 과거 데이터 확인")
    if st.session_state.history:
        for date, records in st.session_state.history.items():
            for hist in records:
                st.dataframe(hist, use_container_width=True)

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        st.session_state.df_raw.to_excel(writer, index=False)
    st.download_button("📥 최종 데이터 다운로드", data=buffer.getvalue(), file_name="결과.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
