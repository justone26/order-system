import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import holidays

# 페이지 설정
st.set_page_config(layout="wide", page_title="재고 관리 시스템")

# 1. 디자인: 제목 및 초기화 (분리형)
st.markdown('<div style="font-size: 55px; font-weight: 900; margin-bottom: 20px;">📦 재고 관리 및 발주 시스템</div>', unsafe_allow_html=True)

if st.button("🔄 시스템 초기화"):
    for key in list(st.session_state.keys()): del st.session_state[key]
    st.rerun()

st.divider()

# 세션 관리
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'history' not in st.session_state: st.session_state.history = {}

# 파일 업로드 (가장 먼저 실행)
uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'csv'])
if uploaded_file:
    df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
    st.session_state.df_raw = df.loc[:, ~df.columns.duplicated()]
    st.rerun()

# 데이터가 있을 때만 각 단계 실행
if st.session_state.df_raw is not None:
    df = st.session_state.df_raw
    cols = df.columns.tolist()

    # 1단계: 매핑 설정
    st.subheader("⚙️ 1단계: 매핑 설정")
    c1, c2 = st.columns(2)
    with c1:
        sold_out = st.selectbox("품절 컬럼", cols, index=0)
        item = st.selectbox("상품명 컬럼", cols, index=0)
    with c2:
        avail = st.selectbox("가용재고 컬럼", cols, index=0)
        option = st.selectbox("옵션 컬럼", cols, index=0)

    # 2~3단계: 분석 설정
    st.subheader("⚙️ 2~3단계: 기간 설정 및 분석")
    if st.button("🚀 분석 실행"):
        df['권장 발주량'] = 0 # 분석 로직
        st.session_state.df_raw = df
        st.success("분석 완료!")
        st.rerun()

    # 4단계: 데이터 편집
    st.subheader("📊 4단계: 데이터 편집")
    edited_df = st.data_editor(df, use_container_width=True)
    if not edited_df.equals(df):
        st.session_state.df_raw = edited_df
        st.rerun()

    # 5단계: 요약
    st.subheader("📋 5단계: 발주 리스트 요약")
    if st.button("💾 기록 저장"):
        date_key = datetime.now().strftime("%Y-%m-%d %H:%M")
        st.session_state.history[date_key] = df.copy()
        st.success("저장 완료!")

    # 6단계: 과거 기록
    st.subheader("📜 6단계: 과거 기록 확인")
    if st.session_state.history:
        date_sel = st.selectbox("날짜 선택", list(st.session_state.history.keys()))
        st.dataframe(st.session_state.history[date_sel], use_container_width=True)
else:
    st.info("파일을 업로드하면 시스템이 활성화됩니다.")
