import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(layout="wide", page_title="재고 관리 시스템")

st.markdown('<div style="font-size: 40px; font-weight: 900;">📦 재고 관리 시스템</div>', unsafe_allow_html=True)

# 1. 파일 업로드 (가장 먼저 위치)
uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'csv'])
if uploaded_file:
    df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
    st.session_state.df_raw = df.loc[:, ~df.columns.duplicated()] # 중복 컬럼 제거
    st.rerun()

# 2. 파일이 있을 때만 단계별 UI 노출
if 'df_raw' in st.session_state and st.session_state.df_raw is not None:
    df = st.session_state.df_raw
    cols = df.columns.tolist()

    # 1단계: 매핑
    st.subheader("⚙️ 1단계: 매핑 설정")
    c1, c2 = st.columns(2)
    with c1:
        sold_out = st.selectbox("품절 컬럼", cols, index=0)
        item = st.selectbox("상품명 컬럼", cols, index=0)
    with c2:
        avail = st.selectbox("가용재고 컬럼", cols, index=0)
        option = st.selectbox("옵션 컬럼", cols, index=0)

    # 강제 컬럼 생성 (에러 방지)
    for col in ['1차 입고예정', '2차 입고예정', '권장 발주량']:
        if col not in df.columns: df[col] = 0

    # 4단계: 데이터 편집 (여기가 제일 중요)
    st.subheader("📊 4단계: 데이터 편집")
    edit_cols = ['상품명', '옵션', '가용재고', '1차 입고예정', '2차 입고예정', '권장 발주량']
    # 실제 존재하는 컬럼만 필터링
    valid_edit_cols = [c for c in edit_cols if c in df.columns]
    
    edited_df = st.data_editor(df[valid_edit_cols], use_container_width=True)
    
    if not edited_df.equals(df[valid_edit_cols]):
        st.session_state.df_raw.update(edited_df)
        st.rerun()

    # 5, 6단계는 여기 아래에 추가하면 됩니다.
    st.success("데이터가 로드되었습니다.")
else:
    st.info("파일을 업로드하면 시스템이 시작됩니다.")
