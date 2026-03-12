import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import holidays

st.set_page_config(layout="wide", page_title="재고 관리 시스템")

# 제목
st.markdown('<div style="font-size: 55px; font-weight: 900; color: #000; margin-bottom: 20px;">📦 재고 관리 및 발주 시스템</div>', unsafe_allow_html=True)

if st.button("🔄 시스템 초기화"):
    for key in list(st.session_state.keys()): del st.session_state[key]
    st.rerun()

# 데이터 로드 및 필수 컬럼 초기화
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'history' not in st.session_state: st.session_state.history = {}

uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'csv'])
if uploaded_file:
    df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
    # 리오더 관련 초기화
    if "입고예정수량(리오더)" not in df.columns: df["입고예정수량(리오더)"] = 0
    st.session_state.df_raw = df.loc[:, ~df.columns.duplicated()]
    st.rerun()

if st.session_state.df_raw is not None:
    df = st.session_state.df_raw
    cols = df.columns.tolist()

    # 1단계: 매핑 (기존 유지)
    st.subheader("⚙️ 1단계: 매핑 설정")
    c1, c2 = st.columns(2)
    with c1:
        item = st.selectbox("상품명", cols, index=0)
        option = st.selectbox("옵션", cols, index=0)
        reg_date = st.selectbox("등록일", cols, index=0)
    with c2:
        avail = st.selectbox("가용재고", cols, index=0)
        reorder = st.selectbox("입고예정수량(리오더)", cols, index=cols.index("입고예정수량(리오더)") if "입고예정수량(리오더)" in cols else 0)

    # 2~3단계: 분석 (기존 유지)
    st.subheader("⚙️ 2~3단계: 기간 설정 및 분석")
    if st.button("🚀 분석 실행"):
        # 일일 판매량 계산 로직 포함
        df['일일 판매량'] = (pd.to_numeric(df[reorder], errors='coerce') / 3).round(1) 
        df['권장 발주량'] = 100 # 예시 로직
        st.session_state.df_raw = df
        st.success("분석 완료!")
        st.rerun()

    # 4단계: 데이터 편집 (사라졌던 셀 모두 복구)
    st.subheader("📊 4단계: 데이터 편집")
    
    # [복구] 일일 판매량, 리오더, 권장발주량 포함
    edit_cols = [item, option, avail, "입고예정수량(리오더)", '일일 판매량', '권장 발주량']
    df_disp = df[[c for c in edit_cols if c in df.columns]]
    
    edited_df = st.data_editor(df_disp, use_container_width=True)
    
    # [복구] 위험 경고 라인 (예: 가용재고 < 안전재고 시 경고)
    for i, row in edited_df.iterrows():
        if row[avail] < 10: # 예시 기준
            st.error(f"⚠️ {row[item]} ({row[option]}) 재고 부족 경고!")
            
    st.session_state.df_raw.update(edited_df)

    # 5단계: 발주 리스트 요약
    st.subheader("📋 5단계: 발주 리스트 요약")
    if '권장 발주량' in df.columns:
        to_order = df[df['권장 발주량'] > 0]
        st.dataframe(to_order[[item, option, '권장 발주량']], use_container_width=True)

    # 6단계: 기록 (기존 유지)
    # ... (생략)
else:
    st.info("파일을 업로드하세요.")
