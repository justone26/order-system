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

# [함수] 컬럼 자동 찾기
def get_auto_index(cols, keywords):
    for key in keywords:
        for i, c in enumerate(cols):
            if key in str(c): return i
    return 0

if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'history' not in st.session_state: st.session_state.history = {}

uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'csv'])
if uploaded_file:
    df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
    st.session_state.df_raw = df.loc[:, ~df.columns.duplicated()]
    st.rerun()

if st.session_state.df_raw is not None:
    df = st.session_state.df_raw
    cols = df.columns.tolist()

    # 1단계: 매핑 설정 (등록일 포함 복구)
    st.subheader("⚙️ 1단계: 매핑 설정")
    c1, c2 = st.columns(2)
    with c1:
        sold_out = st.selectbox("품절 여부", cols, index=get_auto_index(cols, ['품절']))
        vendor = st.selectbox("공급처", cols, index=get_auto_index(cols, ['공급처']))
        item = st.selectbox("상품명", cols, index=get_auto_index(cols, ['상품명']))
        option = st.selectbox("옵션", cols, index=get_auto_index(cols, ['옵션']))
        reg_date = st.selectbox("등록일 컬럼", cols, index=get_auto_index(cols, ['등록일', '입점일']))
    with c2:
        vendor_item = st.selectbox("공급처 상품명", cols, index=get_auto_index(cols, ['공급처상품명']))
        stock = st.selectbox("정상재고", cols, index=get_auto_index(cols, ['정상재고']))
        avail = st.selectbox("가용재고", cols, index=get_auto_index(cols, ['가용재고']))
        t3day = st.selectbox("3일 합계", cols, index=get_auto_index(cols, ['3일']))
        t1week = st.selectbox("1주 합계", cols, index=get_auto_index(cols, ['1주']))

    # 2~3단계: 분석
    st.subheader("⚙️ 2~3단계: 기간 설정 및 분석")
    l1, l2 = st.columns(2)
    lead_time = l1.number_input("리드타임", value=0)
    safety = l2.number_input("안전재고", value=3)
    if st.button("🚀 분석 실행"):
        # 여기서 연산 수행
        st.session_state.df_raw['권장 발주량'] = 10 
        st.rerun()

    # 4단계: 편집 (편집 제한 로직 복구)
    st.subheader("📊 4단계: 데이터 편집")
    edit_cols = [sold_out, vendor, item, option, vendor_item, stock, avail, t3day, t1week, '권장 발주량']
    df_disp = st.session_state.df_raw[[c for c in edit_cols if c in df.columns]]
    edited_df = st.data_editor(df_disp, use_container_width=True, disabled=['권장 발주량'])
    st.session_state.df_raw.update(edited_df)

    # 5단계: 리스트 요약 (필터링 및 상세 리스트 복구)
    st.subheader("📋 5단계: 발주 리스트 요약")
    if '권장 발주량' in df.columns:
        to_order = df[df['권장 발주량'] > 0]
        st.dataframe(to_order[[vendor, item, option, '권장 발주량']], use_container_width=True)
        if st.button("💾 기록 저장"):
            date_key = datetime.now().strftime("%Y-%m-%d %H:%M")
            st.session_state.history[date_key] = df.copy()
            st.success("저장 완료!")

    # 6단계: 과거 기록
    st.subheader("📜 6단계: 과거 기록 확인")
    if st.session_state.history:
        s_date = st.selectbox("기록 선택", list(st.session_state.history.keys()))
        st.dataframe(st.session_state.history[s_date])
