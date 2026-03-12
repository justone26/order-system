import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import holidays

st.set_page_config(layout="wide", page_title="재고 관리 시스템")

# [1] 상태 초기화
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'history' not in st.session_state: st.session_state.history = {}

st.title("📦 재고 관리 및 발주 시스템")

if st.button("🔄 시스템 전체 초기화"):
    for key in list(st.session_state.keys()): del st.session_state[key]
    st.rerun()

# [파일 업로드]
uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'xls', 'csv'])
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
    st.session_state.df_raw = df.loc[:, ~df.columns.duplicated()]
    if "1차 리오더" not in df.columns: st.session_state.df_raw["1차 리오더"] = 0
    if "2차 리오더" not in df.columns: st.session_state.df_raw["2차 리오더"] = 0
    st.success("데이터 로드 완료!")

# [핵심] 데이터가 로드된 후의 전체 로직
if st.session_state.df_raw is not None:
    cols = st.session_state.df_raw.columns.tolist()

    # 1단계: 매핑
    st.subheader("⚙️ 1단계: 자동 매핑 설정")
    def get_auto_index(cols, keywords):
        for key in keywords:
            for i, c in enumerate(cols):
                if key in str(c): return i
        return 0

    c1, c2 = st.columns(2)
    with c1:
        sold_out = st.selectbox("품절 여부", cols, index=get_auto_index(cols, ['품절', '판매중단']))
        vendor = st.selectbox("공급처", cols, index=get_auto_index(cols, ['공급처', '업체명']))
        item = st.selectbox("상품명", cols, index=get_auto_index(cols, ['상품명', '상품']))
        option = st.selectbox("옵션", cols, index=get_auto_index(cols, ['옵션']))
        vendor_item_name = st.selectbox("공급처 상품명", cols, index=get_auto_index(cols, ['공급처상품명', '거래처옵션']))
    with c2:
        stock = st.selectbox("정상재고", cols, index=get_auto_index(cols, ['정상재고', '재고']))
        avail = st.selectbox("가용재고", cols, index=get_auto_index(cols, ['가용재고', '가용']))
        t3day = st.selectbox("3일 발주 합계", cols, index=get_auto_index(cols, ['3일']))
        t1week = st.selectbox("7일 발주 합계", cols, index=get_auto_index(cols, ['7일', '1주']))

    # 2~3단계: 분석 설정
    st.subheader("⚙️ 2~3단계: 분석 설정")
    l1, l2 = st.columns(2)
    lead_time = l1.number_input("리드타임 (일)", value=0)
    safety_stock = l2.number_input("안전재고 (일)", value=3)
    
    if st.button("🚀 분석 실행", type="primary"):
        df = st.session_state.df_raw.copy()
        df['일일 판매량'] = (pd.to_numeric(df[t3day], errors='coerce').fillna(0) / 3).round(0).astype(int)
        df['권장 발주량'] = ((df['일일 판매량'] * (lead_time + safety_stock)) - (pd.to_numeric(df[avail], errors='coerce').fillna(0) + pd.to_numeric(df["1차 리오더"], errors='coerce') + pd.to_numeric(df["2차 리오더"], errors='coerce'))).clip(lower=0).astype(int)
        st.session_state.df_raw = df
        st.rerun()

    # 4단계: 데이터 편집
    st.subheader("📊 4단계: 데이터 편집")
    f1, f2 = st.columns([3, 1])
    search_query = f1.text_input("🔍 상품명 검색")
    filter_mode = f2.selectbox("품절 필터", ["정상만", "품절만", "전체보기"], index=0)
    
    df_filtered = st.session_state.df_raw.copy()
    if filter_mode == "정상만": df_filtered = df_filtered[~df_filtered[sold_out].astype(str).str.contains('품절', na=False)]
    elif filter_mode == "품절만": df_filtered = df_filtered[df_filtered[sold_out].astype(str).str.contains('품절', na=False)]
    if search_query: df_filtered = df_filtered[df_filtered[item].astype(str).str.contains(search_query, na=False)]

    edited_df = st.data_editor(df_filtered, use_container_width=True, key="main_editor")
    st.session_state.df_raw.update(edited_df)

    # 5단계: 발주 리스트 요약
    st.subheader("📋 5단계: 발주 리스트 요약")
    if '권장 발주량' in st.session_state.df_raw.columns:
        to_order = st.session_state.df_raw[st.session_state.df_raw['권장 발주량'] > 0].copy()
        if not to_order.empty:
            summary_cols = [c for c in [vendor, item, option, '권장 발주량'] if c in to_order.columns]
            to_order_summary = to_order[summary_cols]
            st.dataframe(to_order_summary, use_container_width=True)
            
            c1, c2 = st.columns(2)
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine='openpyxl') as w: to_order_summary.to_excel(w, index=False)
            c1.download_button("📥 요약 발주서 다운로드", data=buf.getvalue(), file_name=f"발주요약_{datetime.now().strftime('%m%d')}.xlsx")
            if c2.button("💾 이 리스트 기록 저장"):
                time_key = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                st.session_state.history[time_key] = to_order_summary.copy()
                st.success("기록 저장 완료!")
        else: st.info("발주 대상 없음")

    # 6단계: 과거 기록
    st.subheader("📜 6단계: 과거 데이터 확인")
    if st.session_state.history:
        s_time = st.selectbox("⏰ 저장된 기록 선택", sorted(st.session_state.history.keys(), reverse=True))
        st.dataframe(st.session_state.history[s_time], use_container_width=True)
