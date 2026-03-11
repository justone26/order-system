import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# 1. 페이지 설정 및 세션 초기화
st.set_page_config(layout="wide", page_title="재고 관리 시스템")
st.title("📦 재고 관리 및 발주 시스템")

if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'history' not in st.session_state: st.session_state.history = {}
if 'map_cfg' not in st.session_state:
    st.session_state.map_cfg = {'품절': None, '공급처': None, '상품명': None, '가용재고': None, '3일합계': None}

# 2. 파일 업로드
uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'csv'])
if uploaded_file is not None and st.session_state.df_raw is None:
    st.session_state.df_raw = pd.read_excel(uploaded_file) if uploaded_file.name.endswith('.xlsx') else pd.read_csv(uploaded_file)
    st.session_state.df_raw["입고예정수량(리오더)"] = 0
    st.rerun()

if st.session_state.df_raw is not None:
    cols = st.session_state.df_raw.columns.tolist()

    # [1단계: 컬럼 매핑 - 세션 고정]
    st.subheader("⚙️ 1단계: 컬럼 매핑")
    c1, c2 = st.columns(2)
    with c1:
        st.session_state.map_cfg['품절'] = st.selectbox("품절 여부", cols, index=cols.index(st.session_state.map_cfg['품절']) if st.session_state.map_cfg['품절'] in cols else 0)
        st.session_state.map_cfg['공급처'] = st.selectbox("공급처", cols, index=cols.index(st.session_state.map_cfg['공급처']) if st.session_state.map_cfg['공급처'] in cols else 0)
        st.session_state.map_cfg['상품명'] = st.selectbox("상품명", cols, index=cols.index(st.session_state.map_cfg['상품명']) if st.session_state.map_cfg['상품명'] in cols else 0)
    with c2:
        st.session_state.map_cfg['가용재고'] = st.selectbox("가용재고", cols, index=cols.index(st.session_state.map_cfg['가용재고']) if st.session_state.map_cfg['가용재고'] in cols else 0)
        st.session_state.map_cfg['3일합계'] = st.selectbox("3일 발주 합계", cols, index=cols.index(st.session_state.map_cfg['3일합계']) if st.session_state.map_cfg['3일합계'] in cols else 0)

    # [2~3단계: 분석]
    st.subheader("⚙️ 2~3단계: 분석 실행")
    l1, l2 = st.columns(2)
    lead_time = l1.number_input("리드타임", value=0)
    safety_stock = l2.number_input("안전재고", value=3)
    
    if st.button("🚀 분석 실행"):
        c_avail = st.session_state.map_cfg['가용재고']
        c_3day = st.session_state.map_cfg['3일합계']
        st.session_state.df_raw['권장 발주량'] = (
            (pd.to_numeric(st.session_state.df_raw[c_3day], errors='coerce') / 3) * (lead_time + safety_stock) - 
            (pd.to_numeric(st.session_state.df_raw[c_avail], errors='coerce') + st.session_state.df_raw["입고예정수량(리오더)"])
        ).clip(lower=0).round(0)
        st.success("분석 완료!")

    # [4단계: 매핑 열 편집]
    st.subheader("📊 4단계: 매핑 열 편집")
    target_cols = [c for c in st.session_state.map_cfg.values() if c is not None] + ['입고예정수량(리오더)', '권장 발주량']
    df_disp = st.session_state.df_raw[[c for c in target_cols if c in st.session_state.df_raw.columns]].copy()
    edited_df = st.data_editor(df_disp, use_container_width=True, disabled=[c for c in df_disp.columns if c != "입고예정수량(리오더)"])
    st.session_state.df_raw.update(edited_df)

    # [5단계: 기록 저장]
    st.subheader("📋 5단계: 리스트 기록 저장")
    # 4단계와 동일한 열 구조로 데이터 저장
    df_to_save = st.session_state.df_raw[[c for c in target_cols if c in st.session_state.df_raw.columns]].copy()
    to_order = df_to_save[df_to_save['권장 발주량'] > 0]
    
    if not to_order.empty:
        st.dataframe(to_order, use_container_width=True)
        if st.button("💾 이 리스트 영구 저장"):
            date_key = datetime.now().strftime("%Y-%m-%d")
            record = to_order.copy()
            record['저장시각'] = datetime.now().strftime("%H:%M:%S")
            if date_key not in st.session_state.history: st.session_state.history[date_key] = []
            st.session_state.history[date_key].append(record)
            st.success("저장 완료!")
    
    # [6단계: 과거 확인]
    st.subheader("📜 6단계: 과거 확인")
    if st.session_state.history:
        date_list = sorted(st.session_state.history.keys(), reverse=True)
        selected_date = st.selectbox("날짜 선택", date_list)
        for hist in st.session_state.history[selected_date]:
            with st.expander(f"저장 시각: {hist['저장시각'].iloc[0]}"):
                st.dataframe(hist, use_container_width=True)
