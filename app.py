import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(layout="wide", page_title="재고 관리 시스템")

# 제목 (CSS 분리)
st.markdown('<div style="font-size: 55px; font-weight: 900; margin-bottom: 20px;">📦 재고 관리 및 발주 시스템</div>', unsafe_allow_html=True)

if st.button("🔄 시스템 초기화"):
    for key in list(st.session_state.keys()): del st.session_state[key]
    st.rerun()

# 파일 업로드 로직
uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'csv'])
if uploaded_file:
    df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
    st.session_state.df_raw = df.loc[:, ~df.columns.duplicated()]
    st.rerun()

# 파일이 업로드된 이후의 로직
if 'df_raw' in st.session_state and st.session_state.df_raw is not None:
    df = st.session_state.df_raw
    cols = df.columns.tolist()

    # 1단계: 매핑 설정 (항목 복구)
    st.subheader("⚙️ 1단계: 매핑 설정")
    c1, c2 = st.columns(2)
    with c1:
        sold_out = st.selectbox("품절 여부", cols)
        vendor = st.selectbox("공급처", cols)
        item = st.selectbox("상품명", cols)
        option = st.selectbox("옵션", cols)
    with c2:
        vendor_item = st.selectbox("공급처 상품명", cols)
        stock = st.selectbox("정상재고", cols)
        avail = st.selectbox("가용재고", cols)
        t3day = st.selectbox("3일 판매합계", cols)

    # 2단계~3단계: 분석
    st.subheader("⚙️ 2~3단계: 분석 설정")
    if st.button("🚀 분석 실행"):
        # 분석 로직 예시
        st.session_state.df_raw['권장 발주량'] = 100 
        st.success("✅ 분석 완료!")
        st.rerun()

    # 4단계: 편집
    st.subheader("📊 4단계: 데이터 편집")
    edit_df = st.data_editor(df, use_container_width=True)
    if not edit_df.equals(df):
        st.session_state.df_raw = edit_df
        st.rerun()

 # 5단계: 발주 리스트 요약 (수정된 안전 로직)
    st.subheader("📋 5단계: 발주 리스트 요약")
    if '권장 발주량' in df.columns:
        # 1. 중복 컬럼 제거 (중복된 열이 있으면 첫 번째 열만 남김)
        df = df.loc[:, ~df.columns.duplicated()]
        
        order_list = df[df['권장 발주량'] > 0]
        
        if not order_list.empty:
            # 2. 선택된 컬럼들이 실제 데이터프레임에 있는지 확인 후 추출
            required_cols = [vendor, item, option, '권장 발주량']
            valid_cols = [c for c in required_cols if c in df.columns]
            
            # 3. 중복되지 않은 리스트로 최종 출력
            st.dataframe(order_list[valid_cols].drop_duplicates(), use_container_width=True)
        else:
            st.info("발주할 상품이 없습니다.")
    else:
        st.warning("먼저 '분석 실행'을 진행해주세요.")
