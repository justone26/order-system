import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import holidays

st.set_page_config(layout="wide", page_title="재고 관리 시스템")

# [제목] 클릭 시 새로고침
st.markdown("""
    <style>.title-link { text-decoration: none; color: inherit; }</style>
    <a href="/" class="title-link"><h1>📦 재고 관리 및 발주 시스템</h1></a>
""", unsafe_allow_html=True)

# [세션 관리]
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'history' not in st.session_state: st.session_state.history = {}

# [파일 업로드]
uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'xls', 'csv'])
if uploaded_file is not None and st.session_state.df_raw is None:
    df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
    st.session_state.df_raw = df.loc[:, ~df.columns.duplicated()]
    # 필수 컬럼 생성
    if "입고예정수량(리오더)" not in st.session_state.df_raw.columns:
        st.session_state.df_raw["입고예정수량(리오더)"] = 0
    st.rerun()

if st.session_state.df_raw is not None:
    df = st.session_state.df_raw
    cols = df.columns.tolist()

    # 1단계: 매핑 설정
    st.subheader("⚙️ 1단계: 자동 매핑 설정")
    c1, c2 = st.columns(2)
    with c1:
        item = st.selectbox("상품명", cols, index=0)
        option = st.selectbox("옵션", cols, index=0)
        avail = st.selectbox("가용재고", cols, index=0)
    with c2:
        reorder = st.selectbox("입고예정수량(리오더)", cols, index=cols.index("입고예정수량(리오더)") if "입고예정수량(리오더)" in cols else 0)
        t3day = st.selectbox("3일 발주 합계", cols, index=0)

    # 2~3단계: 분석
    st.subheader("⚙️ 2~3단계: 분석")
    if st.button("🚀 분석 실행"):
        # 일일 판매량 계산 및 권장 발주량 연산
        df['일일 판매량'] = (pd.to_numeric(st.session_state.df_raw[t3day], errors='coerce') / 3).round(1)
        df['권장 발주량'] = (df['일일 판매량'] * 3 - (pd.to_numeric(st.session_state.df_raw[avail], errors='coerce') + pd.to_numeric(st.session_state.df_raw[reorder], errors='coerce'))).clip(lower=0).round(0)
        st.session_state.df_raw = df
        st.success("✅ 분석 완료!")
        st.rerun()

    # 4단계: 편집
    st.subheader("📊 4단계: 데이터 편집")
    edit_cols = [item, option, avail, "입고예정수량(리오더)", '일일 판매량', '권장 발주량']
    df_disp = st.session_state.df_raw[[c for c in edit_cols if c in st.session_state.df_raw.columns]]
    edited_df = st.data_editor(df_disp, use_container_width=True, disabled=['일일 판매량', '권장 발주량'])
    st.session_state.df_raw.update(edited_df)

    # [기능] 위험 경고 (가용재고 5 이하)
    st.subheader("⚠️ 위험 경고")
    danger = st.session_state.df_raw[pd.to_numeric(st.session_state.df_raw[avail], errors='coerce') < 5]
    if not danger.empty:
        st.warning(f"재고 부족 상품 발견: {len(danger)}개")
        st.dataframe(danger[[item, avail]], use_container_width=True)

    # 5단계: 요약
    st.subheader("📋 5단계: 발주 리스트 요약")
    if '권장 발주량' in st.session_state.df_raw.columns:
        to_order = st.session_state.df_raw[st.session_state.df_raw['권장 발주량'] > 0]
        st.dataframe(to_order, use_container_width=True)
        if st.button("💾 기록 저장"):
            date_key = datetime.now().strftime("%Y-%m-%d %H:%M")
            st.session_state.history[date_key] = st.session_state.df_raw.copy()
            st.success("저장 완료!")

    # 6단계: 과거 데이터
    st.subheader("📜 6단계: 과거 기록")
    if st.session_state.history:
        s_date = st.selectbox("기록 선택", list(st.session_state.history.keys()))
        st.dataframe(st.session_state.history[s_date])
