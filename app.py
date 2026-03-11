import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="재고 관리 및 발주 시스템", layout="wide")
st.title("📦 재고 관리 및 발주 시스템")

# [세션 상태 초기화]
if 'df_raw' not in st.session_state:
    st.session_state.df_raw = None

uploaded_file = st.file_uploader("엑셀 또는 CSV 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

if uploaded_file is not None:
    # 파일 로드 및 세션 저장
    if st.session_state.df_raw is None:
        if uploaded_file.name.endswith('.csv'):
            try: st.session_state.df_raw = pd.read_csv(uploaded_file, encoding='utf-8')
            except: st.session_state.df_raw = pd.read_csv(uploaded_file, encoding='cp949')
        else:
            st.session_state.df_raw = pd.read_excel(uploaded_file)
        # 중복 컬럼 제거
        st.session_state.df_raw = st.session_state.df_raw.loc[:, ~st.session_state.df_raw.columns.duplicated()]

if st.session_state.df_raw is not None:
    df = st.session_state.df_raw
    columns = df.columns.tolist()

    # 1. 매핑 섹션 (항상 유지됨)
    st.subheader("⚙️ 1단계: 자동 매핑 확인")
    col1, col2 = st.columns(2)
    with col1:
        sold_out = st.selectbox("품절 여부", columns)
        vendor = st.selectbox("공급처", columns)
        item = st.selectbox("상품명", columns)
        option = st.selectbox("옵션", columns)
    with col2:
        stock = st.selectbox("정상재고", columns)
        avail = st.selectbox("가용재고", columns)
        t3day = st.selectbox("3일 발주 합계", columns)
        t1week = st.selectbox("1주 발주 합계", columns)

    # 2. 리오더 수량 컬럼 확인
    if "입고예정수량(리오더)" not in df.columns:
        df["입고예정수량(리오더)"] = 0

    # 3. 분석 버튼
    if st.button("🚀 분석 실행"):
        df['일일 판매량(기준)'] = (pd.to_numeric(df[t3day], errors='coerce') / 3).fillna(0)
        df['권장 발주량'] = (df['일일 판매량(기준)'] * 3 - (pd.to_numeric(df[avail], errors='coerce') + df["입고예정수량(리오더)"])).clip(lower=0).astype(int)
        st.session_state.df_raw = df # 계산된 결과를 세션에 저장

    # 4. 결과 출력 및 수정 (데이터 에디터)
    st.subheader("📊 분석 결과 및 수량 수정")
    
    edited_df = st.data_editor(st.session_state.df_raw, use_container_width=True)
    
    # 수정된 내용을 세션에 바로 반영
    st.session_state.df_raw = edited_df

    # 5. 다운로드
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        edited_df.to_excel(writer, index=False)
    st.download_button("📥 최종 결과 엑셀 다운로드", data=buffer.getvalue(), file_name="최종_발주서.xlsx")
