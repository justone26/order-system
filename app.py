import streamlit as st
import pandas as pd
from datetime import datetime
import holidays

st.set_page_config(layout="wide", page_title="재고 관리 시스템")

# 제목
st.markdown('<div style="font-size: 55px; font-weight: 900; margin-bottom: 20px;">📦 재고 관리 및 발주 시스템</div>', unsafe_allow_html=True)

if st.button("🔄 시스템 초기화"):
    for key in list(st.session_state.keys()): del st.session_state[key]
    st.rerun()

# [데이터 안전 로드]
uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'csv'])
if uploaded_file:
    df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
    # 중복 컬럼 제거 및 필수 컬럼 강제 생성
    df = df.loc[:, ~df.columns.duplicated()]
    if "입고예정수량(리오더)" not in df.columns:
        df["입고예정수량(리오더)"] = 0
    st.session_state.df_raw = df
    st.rerun()

if st.session_state.df_raw is not None:
    df = st.session_state.df_raw
    cols = df.columns.tolist()

    # 1단계: 매핑 설정
    st.subheader("⚙️ 1단계: 매핑 설정")
    c1, c2 = st.columns(2)
    with c1:
        sold_out = st.selectbox("품절 여부", cols, index=0)
        item = st.selectbox("상품명", cols, index=0)
    with c2:
        avail = st.selectbox("가용재고", cols, index=0)
        reorder = st.selectbox("입고예정수량(리오더)", cols, index=cols.index("입고예정수량(리오더)") if "입고예정수량(리오더)" in cols else 0)

    # 2단계: 분석 (컬럼 체크 로직 강화)
    st.subheader("⚙️ 2~3단계: 분석")
    if st.button("🚀 분석 실행"):
        # 분석 시 컬럼 존재 여부 재확인
        if reorder in df.columns:
            df['권장 발주량'] = (100 - pd.to_numeric(df[avail], errors='coerce').fillna(0) - pd.to_numeric(df[reorder], errors='coerce').fillna(0)).clip(lower=0)
        else:
            df['권장 발주량'] = 0
        st.session_state.df_raw = df
        st.success("분석 완료!")
        st.rerun()

    # 4단계: 편집
    st.subheader("📊 4단계: 데이터 편집")
    edited_df = st.data_editor(df, use_container_width=True)
    if not edited_df.equals(df):
        st.session_state.df_raw = edited_df
        st.rerun()

    # 5단계: 요약
    st.subheader("📋 5단계: 발주 리스트 요약")
    if '권장 발주량' in df.columns:
        st.dataframe(df[df['권장 발주량'] > 0], use_container_width=True)
else:
    st.info("파일을 업로드하여 시스템을 시작하세요.")
