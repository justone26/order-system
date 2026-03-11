import streamlit as st
import pandas as pd
import sqlite3
from io import BytesIO
from datetime import datetime

# DB 초기화
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

# [세션 상태 관리]
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'history' not in st.session_state: st.session_state.history = {}

def get_idx(cols, keywords):
    for key in keywords:
        for i, c in enumerate(cols):
            if key in str(c): return i
    return 0

# 1. 파일 업로드
uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'xls', 'csv'])
if uploaded_file is not None and st.session_state.df_raw is None:
    df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
    st.session_state.df_raw = df.loc[:, ~df.columns.duplicated()]
    if "입고예정수량(리오더)" not in st.session_state.df_raw.columns:
        st.session_state.df_raw["입고예정수량(리오더)"] = 0
    st.rerun()

if st.session_state.df_raw is not None:
    cols = st.session_state.df_raw.columns.tolist()

    # [1단계: 매핑 설정]
    st.subheader("⚙️ 1단계: 자동 매핑 설정")
    c1, c2 = st.columns(2)
    with c1:
        sold_out = st.selectbox("품절 여부", cols, index=get_idx(cols, ['품절', '판매중단']))
        vendor = st.selectbox("공급처", cols, index=get_idx(cols, ['공급처', '업체명']))
        item = st.selectbox("상품명", cols, index=get_idx(cols, ['상품명', '상품']))
    with c2:
        avail = st.selectbox("가용재고", cols, index=get_idx(cols, ['가용재고', '가용']))
        t3day = st.selectbox("3일 발주 합계", cols, index=get_idx(cols, ['3일', '최근3일']))

    # [2단계: 기간 설정]
    st.subheader("⚙️ 2단계: 기간 설정")
    l1, l2 = st.columns(2)
    lead_time = l1.number_input("리드타임 (일)", value=0)
    safety_stock = l2.number_input("안전재고 (일)", value=3)

    # [3단계: 분석 실행]
    if st.button("🚀 분석 실행"):
        st.session_state.df_raw['일일 판매량'] = (pd.to_numeric(st.session_state.df_raw[t3day], errors='coerce') / 3).round(0)
        st.session_state.df_raw['권장 발주량'] = (st.session_state.df_raw['일일 판매량'] * (lead_time + safety_stock) - 
                                            (pd.to_numeric(st.session_state.df_raw[avail], errors='coerce') + st.session_state.df_raw["입고예정수량(리오더)"])).clip(lower=0)
        st.rerun()

    # [4단계: 검색 및 데이터 편집]
    st.subheader("📊 4단계: 검색 및 데이터 편집")
    f1, f2 = st.columns([2, 1])
    search = f1.text_input("🔍 상품명 검색")
    status_filter = f2.selectbox("🚫 품절 필터", ["전체보기"] + st.session_state.df_raw[sold_out].unique().tolist())
    
    df_disp = st.session_state.df_raw.copy()
    if status_filter != "전체보기": df_disp = df_disp[df_disp[sold_out] == status_filter]
    if search: df_disp = df_disp[df_disp[item].astype(str).str.contains(search, na=False)]

    edited_df = st.data_editor(df_disp, use_container_width=True)
    st.session_state.df_raw.update(edited_df)

    # [5단계: 발주 요약 및 저장]
    if st.button("💾 리스트 영구 저장"):
        conn = sqlite3.connect('inventory.db')
        to_save = st.session_state.df_raw[st.session_state.df_raw['권장 발주량'] > 0].copy()
        to_save['date'] = datetime.now().strftime("%Y-%m-%d")
        to_save.to_sql('history', conn, if_exists='append', index=False)
        conn.close()
        st.success("데이터베이스에 저장되었습니다!")

    # [6단계: 다운로드]
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        st.session_state.df_raw.to_excel(writer, index=False)
    st.download_button("📥 최종 데이터 다운로드", data=buffer.getvalue(), file_name="결과.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
