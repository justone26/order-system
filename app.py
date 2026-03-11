import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(layout="wide", page_title="재고 관리 시스템")
st.title("📦 재고 관리 및 발주 시스템")

if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'history' not in st.session_state: st.session_state.history = {}

def get_idx(cols, keywords):
    for key in keywords:
        for i, c in enumerate(cols):
            if key in str(c): return i
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

    # [1단계: 복구된 풍성한 매핑]
    st.subheader("⚙️ 1단계: 자동 매핑 설정")
    c1, c2 = st.columns(2)
    with c1:
        sold_out = st.selectbox("품절 여부", cols, index=get_idx(cols, ['품절', '판매중단']))
        vendor = st.selectbox("공급처", cols, index=get_idx(cols, ['공급처', '업체명']))
        item = st.selectbox("상품명", cols, index=get_idx(cols, ['상품명', '상품']))
        option = st.selectbox("옵션", cols, index=get_idx(cols, ['옵션']))
        vendor_opt = st.selectbox("공급처옵션", cols, index=get_idx(cols, ['공급처옵션', '거래처옵션']))
    with c2:
        stock = st.selectbox("정상재고", cols, index=get_idx(cols, ['정상재고', '재고']))
        avail = st.selectbox("가용재고", cols, index=get_idx(cols, ['가용재고', '가용']))
        t3day = st.selectbox("3일 발주 합계", cols, index=get_idx(cols, ['3일', '최근3일']))
        t1week = st.selectbox("1주 발주 합계", cols, index=get_idx(cols, ['1주', '7일', '최근7일']))

    # [2~3단계: 설정 및 분석]
    st.subheader("⚙️ 2단계: 기간 및 분석")
    l1, l2 = st.columns(2)
    lead_time = l1.number_input("리드타임", value=0)
    safety_stock = l2.number_input("안전재고", value=3)
    if st.button("🚀 분석 실행"):
        st.session_state.df_raw['일일 판매량'] = (pd.to_numeric(st.session_state.df_raw[t3day], errors='coerce') / 3).round(0)
        st.session_state.df_raw['권장 발주량'] = (st.session_state.df_raw['일일 판매량'] * (lead_time + safety_stock) - 
                                            (pd.to_numeric(st.session_state.df_raw[avail], errors='coerce') + st.session_state.df_raw["입고예정수량(리오더)"])).clip(lower=0)
        st.rerun()

    # [4단계: 데이터 편집 (선택한 9개 항목만 표시)]
    st.subheader("📊 4단계: 데이터 편집")
    edit_cols = [sold_out, vendor, item, option, vendor_opt, stock, avail, "입고예정수량(리오더)", t3day, t1week, '권장 발주량']
    # 실제 존재하는 컬럼만 필터링
    df_final = st.session_state.df_raw[[c for c in edit_cols if c in st.session_state.df_raw.columns]].copy()
    
    edited_df = st.data_editor(df_final, use_container_width=True, 
                               disabled=[c for c in df_final.columns if c != "입고예정수량(리오더)"])
    st.session_state.df_raw.update(edited_df)

    # [5단계: 발주 필요 요약]
    st.subheader("📋 5단계: 발주 필요 리스트 요약")
    to_order = st.session_state.df_raw[st.session_state.df_raw['권장 발주량'] > 0]
    st.dataframe(to_order[edit_cols], use_container_width=True)
    
    if st.button("💾 리스트 기록 저장"):
        date_key = datetime.now().strftime("%Y-%m-%d")
        record = to_order[edit_cols].copy()
        record['저장시각'] = datetime.now().strftime("%H:%M:%S")
        if date_key not in st.session_state.history: st.session_state.history[date_key] = []
        st.session_state.history[date_key].append(record)
        st.success("저장 완료!")

    # [6단계: 과거 기록 조회]
    st.subheader("📜 6단계: 과거 데이터 확인")
    if st.session_state.history:
        selected_date = st.selectbox("조회할 날짜 선택", sorted(st.session_state.history.keys(), reverse=True))
        for hist in st.session_state.history[selected_date]:
            st.dataframe(hist, use_container_width=True)
