import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import holidays

st.set_page_config(layout="wide", page_title="재고 관리 시스템")

# [1] 상태 초기화: 키가 없으면 생성
if 'file_key' not in st.session_state:
    st.session_state.file_key = 0
if 'df_raw' not in st.session_state:
    st.session_state.df_raw = None

st.title("📦 재고 관리 및 발주 시스템")

# [2] 초기화 버튼: file_key 값을 변경하여 위젯을 강제로 새로고침
if st.button("🔄 시스템 초기화 (데이터 및 파일 삭제)"):
    # 세션 내 모든 데이터 제거
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    # 초기화 후 필수 키만 다시 생성
    st.session_state.file_key = 1 
    st.rerun()

# [3] 파일 업로더: key값을 사용하여 상태 변경 시 위젯이 사라졌다 다시 생성됨
uploaded_file = st.file_uploader(
    "엑셀/CSV 업로드", 
    type=['xlsx', 'xls', 'csv'], 
    key=f"uploader_{st.session_state.file_key}"
)

# [4] 데이터 로드 로직
if uploaded_file is not None and st.session_state.df_raw is None:
    try:
        df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
        st.session_state.df_raw = df.loc[:, ~df.columns.duplicated()]
        if "입고예정수량(리오더)" not in st.session_state.df_raw.columns:
            st.session_state.df_raw["입고예정수량(리오더)"] = 0
        st.rerun()
    except Exception as e:
        st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")

# [5] 데이터 처리 영역 (이 아래에 1~6단계 로직을 붙여넣어)
if st.session_state.df_raw is not None:
    st.success("데이터가 로드되었습니다!")
    # 여기에 너의 기존 1단계 매핑 로직부터 붙이면 돼.
