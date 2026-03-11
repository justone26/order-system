import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(layout="wide")
st.title("📦 재고 관리 및 발주 시스템")

# [세션 관리]
if 'df_raw' not in st.session_state: st.session_state.df_raw = None

# [헬퍼 함수]
def get_idx(col_list, keywords):
    for key in keywords:
        for i, c in enumerate(col_list):
            if key in c: return i
    return 0

# 1. 파일 업로드
uploaded_file = st.file_uploader("엑셀 또는 CSV 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

if uploaded_file is not None and st.session_state.df_raw is None:
    df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
    st.session_state.df_raw = df.loc[:, ~df.columns.duplicated()]
    st.session_state.df_raw["입고예정수량(리오더)"] = 0
    st.rerun()

if st.session_state.df_raw is not None:
    cols = st.session_state.df_raw.columns.tolist()

    # 1단계: 설정
    st.subheader("⚙️ 1단계: 매핑 및 기간 설정")
    with st.expander("필수 설정", expanded=True):
        c1, c2 = st.columns(2)
        with c1:
            sold_out = st.selectbox("품절 여부", cols, index=get_idx(cols, ['품절', '판매중단']))
            item = st.selectbox("상품명", cols, index=get_idx(cols, ['상품명', '상품']))
            avail = st.selectbox("가용재고", cols, index=get_idx(cols, ['가용재고', '가용']))
        with c2:
            t3day = st.selectbox("3일 판매 합계", cols, index=get_idx(cols, ['3일', '최근3일']))
            lead_time = st.number_input("리드타임 (일)", value=0)
            safety = st.number_input("안전재고 (일)", value=3)

    # 분석 실행
    if st.button("🚀 분석 실행 (계산 반영)"):
        st.session_state.df_raw['일일 판매량'] = (pd.to_numeric(st.session_state.df_raw[t3day], errors='coerce') / 3).round(0)
        st.session_state.df_raw['권장 발주량'] = (st.session_state.df_raw['일일 판매량'] * (lead_time + safety) - 
                                            (pd.to_numeric(st.session_state.df_raw[avail], errors='coerce') + st.session_state.df_raw["입고예정수량(리오더)"])).clip(lower=0)
        st.rerun()

    # 2단계: 필터링 및 편집
    st.subheader("📊 2단계: 검색 및 데이터 편집")
    f1, f2 = st.columns([2, 1])
    search = f1.text_input("🔍 상품명 검색")
    filter_mode = f2.selectbox("품절 필터", ["전체보기", "품절만", "정상만"])
    
    # 보여주기용 데이터 필터링
    df_disp = st.session_state.df_raw.copy()
    if filter_mode == "품절만": df_disp = df_disp[df_disp[sold_out].astype(str).str.upper() == 'Y']
    if filter_mode == "정상만": df_disp = df_disp[df_disp[sold_out].astype(str).str.upper() != 'Y']
    if search: df_disp = df_disp[df_disp[item].astype(str).str.contains(search, na=False)]

    
    
    # 핵심 데이터만 편집기 출력
    display_cols = [sold_out, item, avail, "입고예정수량(리오더)", t3day, '일일 판매량', '권장 발주량']
    df_final = df_disp[[c for c in display_cols if c in df_disp.columns]]
    
    edited_df = st.data_editor(df_final, use_container_width=True)
    
    # 수정 내용 동기화
    st.session_state.df_raw.update(edited_df)

    # 3단계: 다운로드
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        st.session_state.df_raw.to_excel(writer, index=False)
    st.download_button("📥 최종 데이터 다운로드", buffer.getvalue(), "결과.xlsx")
