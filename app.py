import streamlit as st
import pandas as pd
import sqlite3
from io import BytesIO
from datetime import datetime

# 1. DB 초기화
def init_db():
    conn = sqlite3.connect('inventory.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS history 
                 (date TEXT, vendor TEXT, item TEXT, qty INTEGER, time TEXT)''')
    conn.commit()
    conn.close()

init_db()

st.set_page_config(layout="wide", page_title="재고 관리 시스템")
st.title("📦 재고 관리 및 발주 시스템")

if 'df_raw' not in st.session_state: st.session_state.df_raw = None

uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'csv'])
if uploaded_file is not None and st.session_state.df_raw is None:
    st.session_state.df_raw = pd.read_excel(uploaded_file) if uploaded_file.name.endswith('.xlsx') else pd.read_csv(uploaded_file)
    st.rerun()

if st.session_state.df_raw is not None:
    cols = st.session_state.df_raw.columns.tolist()

    # [1단계: 모든 컬럼 매핑]
    st.subheader("⚙️ 1단계: 전체 컬럼 매핑")
    c1, c2 = st.columns(2)
    with c1:
        sold_out = st.selectbox("품절 여부 컬럼", cols)
        vendor = st.selectbox("공급처 컬럼", cols)
        item = st.selectbox("상품명 컬럼", cols)
        option = st.selectbox("옵션 컬럼", cols)
    with c2:
        stock = st.selectbox("정상재고 컬럼", cols)
        avail = st.selectbox("가용재고 컬럼", cols)
        t3day = st.selectbox("3일 발주 합계 컬럼", cols)
        t1week = st.selectbox("1주 발주 합계 컬럼", cols)

    # [2단계: 기간 설정]
    st.subheader("⚙️ 2단계: 기간 설정")
    l1, l2 = st.columns(2)
    lead_time = l1.number_input("리드타임 (일)", value=0)
    safety_stock = l2.number_input("안전재고 (일)", value=3)

    # [3단계: 분석 실행]
    st.subheader("⚙️ 3단계: 분석 실행")
    if st.button("🚀 분석 실행"):
        st.session_state.df_raw['일일 판매량'] = (pd.to_numeric(st.session_state.df_raw[t3day], errors='coerce') / 3).round(0)
        st.session_state.df_raw['권장 발주량'] = (st.session_state.df_raw['일일 판매량'] * (lead_time + safety_stock) - 
                                            pd.to_numeric(st.session_state.df_raw[avail], errors='coerce')).clip(lower=0)
        st.session_state.df_raw['상태'] = st.session_state.df_raw[sold_out].astype(str)
        st.success("분석 완료!")

    # [4단계: 데이터 편집 및 필터링]
    st.subheader("📊 4단계: 데이터 편집")
    if '권장 발주량' in st.session_state.df_raw.columns:
        f1, f2 = st.columns([2, 1])
        search = f1.text_input("🔍 상품명 검색")
        
        # 품절 필터: 상태 컬럼의 고유값들을 활용
        status_options = ["전체보기"] + st.session_state.df_raw['상태'].unique().tolist()
        filter_mode = f2.selectbox("🚫 품절 필터", status_options)
        
        df_disp = st.session_state.df_raw.copy()
        if filter_mode != "전체보기":
            df_disp = df_disp[df_disp['상태'] == filter_mode]
        if search:
            df_disp = df_disp[df_disp[item].astype(str).str.contains(search, na=False)]
        
        # 편집 가능한 컬럼만 노출
        edited_df = st.data_editor(df_disp, use_container_width=True)
        st.session_state.df_raw.update(edited_df)
    else:
        st.info("3단계 분석을 먼저 실행해주세요.")
