import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(layout="wide", page_title="재고 관리 시스템")

# [1] 상태 초기화
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'history' not in st.session_state: st.session_state.history = {}

st.title("📦 재고 관리 및 발주 시스템")

# 시스템 전체 초기화
if st.button("🔄 시스템 전체 초기화"):
    for key in list(st.session_state.keys()): del st.session_state[key]
    st.rerun()

# [1단계: 업로드]
uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'xls', 'csv'])
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
    st.session_state.df_raw = df.loc[:, ~df.columns.duplicated()]
    if "1차 리오더" not in st.session_state.df_raw.columns: st.session_state.df_raw["1차 리오더"] = 0
    if "2차 리오더" not in st.session_state.df_raw.columns: st.session_state.df_raw["2차 리오더"] = 0
    st.rerun()

# [데이터가 있을 때만 실행]
if st.session_state.df_raw is not None:
    cols = st.session_state.df_raw.columns.tolist()

    # 1단계: 매핑 (등록일 포함 완벽 복구)
    st.subheader("⚙️ 1단계: 매핑 설정")
    c1, c2 = st.columns(2)
    with c1:
        sold_out = st.selectbox("품절 여부", cols, index=0)
        vendor = st.selectbox("공급처", cols, index=1)
        item = st.selectbox("상품명", cols, index=2)
        option = st.selectbox("옵션", cols, index=3)
    with c2:
        reg_date = st.selectbox("등록일 컬럼", cols, index=4) # 등록일 복구
        stock = st.selectbox("정상재고", cols, index=5)
        avail = st.selectbox("가용재고", cols, index=6)
        t3day = st.selectbox("3일 발주 합계", cols, index=7)

    # 2~3단계: 분석
    st.subheader("⚙️ 2~3단계: 분석 설정")
    l1, l2 = st.columns(2)
    lead_time = l1.number_input("리드타임 (일)", value=0)
    safety_stock = l2.number_input("안전재고 (일)", value=3)
    
    if st.button("🚀 분석 실행"):
        df = st.session_state.df_raw.copy()
        df['일일 판매량'] = (pd.to_numeric(df[t3day], errors='coerce').fillna(0) / 3).round(0).astype(int)
        df['권장 발주량'] = ((df['일일 판매량'] * (lead_time + safety_stock)) - (pd.to_numeric(df[avail], errors='coerce') + pd.to_numeric(df["1차 리오더"], errors='coerce') + pd.to_numeric(df["2차 리오더"], errors='coerce'))).clip(lower=0).astype(int)
        st.session_state.df_raw = df
        st.rerun()

    # 4단계: 데이터 편집 (모든 컬럼 포함)
    st.subheader("📊 4단계: 데이터 편집")
    edited_df = st.data_editor(st.session_state.df_raw, use_container_width=True)
    st.session_state.df_raw.update(edited_df)

    # 5단계: 요약 (다운로드 및 저장 복구)
    st.subheader("📋 5단계: 발주 리스트 요약")
    to_order = st.session_state.df_raw[st.session_state.df_raw['권장 발주량'] > 0] if '권장 발주량' in st.session_state.df_raw.columns else pd.DataFrame()
    
    if not to_order.empty:
        st.dataframe(to_order, use_container_width=True)
        c1, c2 = st.columns(2)
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w: to_order.to_excel(w, index=False)
        c1.download_button("📥 요약 발주서 다운로드", data=buf.getvalue(), file_name="발주서.xlsx")
        if c2.button("💾 기록 저장"):
            time_key = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            st.session_state.history[time_key] = to_order.copy()
            st.success("저장 완료!")
    else: st.info("발주 대상 없음")

    # 6단계: 과거 기록 (복구 완료)
    st.subheader("📜 6단계: 과거 데이터 확인")
    if st.session_state.history:
        s_time = st.selectbox("⏰ 저장 기록 선택", sorted(st.session_state.history.keys(), reverse=True))
        st.dataframe(st.session_state.history[s_time], use_container_width=True)
