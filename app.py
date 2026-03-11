import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(layout="wide", page_title="재고 관리 시스템")
st.title("📦 재고 관리 및 발주 시스템")

if 'df_raw' not in st.session_state: st.session_state.df_raw = None

def get_idx(cols, keywords):
    for key in keywords:
        for i, c in enumerate(cols):
            if key in str(c): return i
    return 0

# 1. 업로드 로직
uploaded_file = st.file_uploader("엑셀 또는 CSV 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])
if uploaded_file is not None and st.session_state.df_raw is None:
    df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
    st.session_state.df_raw = df.loc[:, ~df.columns.duplicated()]
    if "입고예정수량(리오더)" not in st.session_state.df_raw.columns:
        st.session_state.df_raw["입고예정수량(리오더)"] = 0
    st.rerun()

if st.session_state.df_raw is not None:
    cols = st.session_state.df_raw.columns.tolist()

    # [1~3단계: 매핑, 설정, 분석] (이전 코드와 동일)
    st.subheader("⚙️ 1단계: 자동 매핑 설정")
    c1, c2 = st.columns(2)
    with c1:
        sold_out = st.selectbox("품절 여부", cols, index=get_idx(cols, ['품절', '판매중단']))
        vendor = st.selectbox("공급처", cols, index=get_idx(cols, ['공급처', '업체명']))
        item = st.selectbox("상품명", cols, index=get_idx(cols, ['상품명', '상품']))
    with c2:
        avail = st.selectbox("가용재고", cols, index=get_idx(cols, ['가용재고', '가용']))
        t3day = st.selectbox("3일 발주 합계", cols, index=get_idx(cols, ['3일', '최근3일']))

    st.subheader("⚙️ 2단계: 기간 설정")
    l1, l2 = st.columns(2)
    lead_time = l1.number_input("리드타임 (일)", value=0)
    safety_stock = l2.number_input("안전재고 (일)", value=3)

    st.subheader("⚙️ 3단계: 분석 실행")
    if st.button("🚀 분석 실행", use_container_width=True):
        st.session_state.df_raw['일일 판매량'] = (pd.to_numeric(st.session_state.df_raw[t3day], errors='coerce') / 3).round(0)
        st.session_state.df_raw['권장 발주량'] = (st.session_state.df_raw['일일 판매량'] * (lead_time + safety_stock) - 
                                            (pd.to_numeric(st.session_state.df_raw[avail], errors='coerce') + st.session_state.df_raw["입고예정수량(리오더)"])).clip(lower=0)
        st.rerun()

    # [4단계: 검색 및 데이터 편집 (알림 기능 포함)]
    st.subheader("📊 4단계: 검색 및 데이터 편집")
    
    # ... (필터링 로직 동일)
    df_disp = st.session_state.df_raw.copy()
    
    # 편집기 구성 및 알림 적용
    edited_df = st.data_editor(
        df_disp, 
        use_container_width=True,
        column_config={
            "권장 발주량": st.column_config.NumberColumn(
                "권장 발주량",
                help="발주량이 0보다 크면 알림",
                format="%d",
            ),
            "입고예정수량(리오더)": st.column_config.NumberColumn("입고예정수량(리오더)", disabled=False)
        },
        disabled=[c for c in df_disp.columns if c != "입고예정수량(리오더)"]
    )
    
    # 발주 필요 상품 알림
    to_order = st.session_state.df_raw[st.session_state.df_raw['권장 발주량'] > 0]
    if not to_order.empty:
        st.warning(f"⚠️ 총 {len(to_order)}개의 상품에 대해 발주가 필요합니다!")
        st.dataframe(to_order[['상품명', '권장 발주량']], use_container_width=True)

    st.session_state.df_raw.update(edited_df)
    
    # 다운로드
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        st.session_state.df_raw.to_excel(writer, index=False)
    st.download_button("📥 최종 결과 다운로드", buffer.getvalue(), "결과.xlsx")
