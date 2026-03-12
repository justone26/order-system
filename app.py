import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import holidays

st.set_page_config(layout="wide", page_title="재고 관리 시스템")

# [1] 상태 및 히스토리 초기화
if 'file_key' not in st.session_state: st.session_state.file_key = 0
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'history' not in st.session_state: st.session_state.history = {}

st.title("📦 재고 관리 및 발주 시스템")

if st.button("🔄 시스템 전체 초기화"):
    for key in list(st.session_state.keys()): del st.session_state[key]
    st.session_state.file_key = 1
    st.rerun()

uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'xls', 'csv'], key=f"uploader_{st.session_state.file_key}")

if uploaded_file is not None and st.session_state.df_raw is None:
    df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
    df = df.loc[:, ~df.columns.duplicated()]
    df["1차 리오더"] = 0
    df["2차 리오더"] = 0
    st.session_state.df_raw = df
    st.rerun()

# [데이터가 로드된 경우 전체 프로세스 실행]
if st.session_state.df_raw is not None:
    cols = st.session_state.df_raw.columns.tolist()

    # 1단계: 매핑
    st.subheader("⚙️ 1단계: 자동 매핑 설정")
    c1, c2 = st.columns(2)
    sold_out = c1.selectbox("품절 여부", cols, index=0)
    vendor = c1.selectbox("공급처", cols, index=0)
    item = c1.selectbox("상품명", cols, index=0)
    option = c1.selectbox("옵션", cols, index=0)
    reg_date_col = c2.selectbox("등록일 컬럼", cols, index=0)
    avail = c2.selectbox("가용재고", cols, index=0)
    t3day = c2.selectbox("3일 발주 합계", cols, index=0)

    # 2~3단계: 분석
    st.subheader("⚙️ 2~3단계: 분석 설정")
    l1, l2 = st.columns(2)
    lead_time = l1.number_input("리드타임 (일)", value=0)
    safety_stock = l2.number_input("안전재고 (일)", value=3)
    
    if st.button("🚀 분석 실행"):
        df = st.session_state.df_raw
        df['일일 판매량'] = (pd.to_numeric(df[t3day], errors='coerce').fillna(0) / 3).round(0).astype(int)
        v_reorder = pd.to_numeric(df["1차 리오더"], errors='coerce').fillna(0) + pd.to_numeric(df["2차 리오더"], errors='coerce').fillna(0)
        df['권장 발주량'] = ((df['일일 판매량'] * (lead_time + safety_stock)) - (pd.to_numeric(df[avail], errors='coerce') + v_reorder)).clip(lower=0).astype(int)
        st.session_state.df_raw = df
        st.rerun()

    # 4단계: 편집
    st.subheader("📊 4단계: 데이터 편집 (1차/2차 리오더 수정 가능)")
    df_disp = st.session_state.df_raw
    edited_df = st.data_editor(
        df_disp, 
        use_container_width=True, 
        disabled=[c for c in df_disp.columns if c not in ["1차 리오더", "2차 리오더"]]
    )
    st.session_state.df_raw.update(edited_df)

    # 5단계: 요약
    st.subheader("📋 5단계: 발주 리스트 요약")
    to_order = st.session_state.df_raw[st.session_state.df_raw['권장 발주량'] > 0]
    if not to_order.empty:
        st.dataframe(to_order, use_container_width=True)
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w: to_order.to_excel(w, index=False)
        st.download_button("📥 발주 리스트 다운로드", data=buf.getvalue(), file_name="발주서.xlsx")
        
        if st.button("💾 기록 저장"):
            st.session_state.history[datetime.now().strftime("%Y-%m-%d %H:%M:%S")] = to_order.copy().reset_index(drop=True)
            st.success("기록 저장 완료!")

    # 6단계: 과거 확인
    st.subheader("📜 6단계: 과거 데이터 확인")
    if st.session_state.history:
        s_time = st.selectbox("⏰ 저장 기록 선택", sorted(st.session_state.history.keys(), reverse=True))
        st.dataframe(st.session_state.history[s_time], use_container_width=True)
