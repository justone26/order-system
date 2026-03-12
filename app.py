import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(layout="wide", page_title="재고 관리 시스템")

# [1] 상태 관리 (데이터를 한 번만 로드)
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'history' not in st.session_state: st.session_state.history = {}

st.title("📦 재고 관리 및 발주 시스템")

# [파일 업로드] - 가장 처음만 동작
uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'xls', 'csv'])
if uploaded_file is not None and st.session_state.df_raw is None:
    df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
    st.session_state.df_raw = df.loc[:, ~df.columns.duplicated()]
    # 초기값 설정
    for col in ["1차 리오더", "2차 리오더"]:
        if col not in st.session_state.df_raw.columns: st.session_state.df_raw[col] = 0
    st.rerun()

# [메인 로직]
if st.session_state.df_raw is not None:
    # 1단계: 매핑 (간소화해서 속도 향상)
    st.subheader("⚙️ 1단계: 매핑 설정")
    cols = st.session_state.df_raw.columns.tolist()
    c1, c2 = st.columns(2)
    sold_out = c1.selectbox("품절 여부", cols, index=0)
    vendor = c1.selectbox("공급처", cols, index=1)
    item = c1.selectbox("상품명", cols, index=2)
    reg_date = c2.selectbox("등록일", cols, index=4)
    stock = c2.selectbox("정상재고", cols, index=5)
    t3day = c2.selectbox("3일 발주합계", cols, index=7)

    # 2~3단계: 분석
    st.subheader("⚙️ 2~3단계: 분석 설정")
    lead_time = st.number_input("리드타임", value=0)
    
    if st.button("🚀 분석 실행"):
        df = st.session_state.df_raw.copy()
        df['일일 판매량'] = (pd.to_numeric(df[t3day], errors='coerce').fillna(0) / 3).round(0)
        df['권장 발주량'] = (df['일일 판매량'] * lead_time).clip(lower=0)
        st.session_state.df_raw = df
        st.rerun() # 여기서만 새로고침

    # 4단계: 편집
    st.subheader("📊 4단계: 데이터 편집")
    edited_df = st.data_editor(st.session_state.df_raw, use_container_width=True)
    st.session_state.df_raw.update(edited_df)

    # 5단계: 저장
    if st.button("💾 발주 리스트 저장"):
        to_order = st.session_state.df_raw[st.session_state.df_raw['권장 발주량'] > 0]
        st.session_state.history[datetime.now().strftime("%H:%M:%S")] = to_order
        st.success("저장 완료!")

    # 6단계: 기록
    st.subheader("📜 6단계: 과거 데이터")
    if st.session_state.history:
        st.selectbox("기록 선택", list(st.session_state.history.keys()))
