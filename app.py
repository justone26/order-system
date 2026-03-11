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
st.subheader("⚙️ 1단계: 자동 매핑 설정")
c1, c2 = st.columns(2)
with c1:
sold_out = st.selectbox("품절 여부", cols, index=get_idx(cols, ['품절']))
vendor = st.selectbox("공급처", cols, index=get_idx(cols, ['공급처']))
item = st.selectbox("상품명", cols, index=get_idx(cols, ['상품명']))
with c2:
avail = st.selectbox("가용재고", cols, index=get_idx(cols, ['가용재고']))
t3day = st.selectbox("3일 발주 합계", cols, index=get_idx(cols, ['3일']))
