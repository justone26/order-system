import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import holidays

st.set_page_config(layout="wide", page_title="재고 관리 시스템")

st.title("📦 재고 관리 및 발주 시스템")

# [1] 세션 및 키 관리
if 'file_key' not in st.session_state: st.session_state.file_key = 0
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'history' not in st.session_state: st.session_state.history = {}

# [2] 초기화 버튼 (눌리면 파일 위젯까지 초기화됨)
if st.button("🔄 시스템 초기화 (데이터 삭제)"):
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.session_state.file_key += 1
    st.rerun()

# [3] 파일 업로더 (key 값을 통해 강제 새로고침 가능)
uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'xls', 'csv'], key=f"up_{st.session_state.file_key}")

if uploaded_file is not None and st.session_state.df_raw is None:
    df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
    st.session_state.df_raw = df.loc[:, ~df.columns.duplicated()]
    if "입고예정수량(리오더)" not in st.session_state.df_raw.columns:
        st.session_state.df_raw["입고예정수량(리오더)"] = 0
    st.rerun()

# [4] 이후 매핑, 분석, 편집 로직들 (이어서 작성)
if st.session_state.df_raw is not None:
    st.success("데이터가 로드되었습니다!")
    # 여기에 아까 작성한 로직들을 붙여넣으면 돼!
