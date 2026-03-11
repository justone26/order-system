import streamlit as st
import pandas as pd
from io import BytesIO

# 페이지 설정
st.set_page_config(page_title="재고 관리 및 발주 시스템", layout="wide")
st.title("📦 재고 관리 및 발주 시스템")

# 1. 세션 상태 초기화
if 'df_data' not in st.session_state:
    st.session_state.df_data = None

def get_best_match(keywords, cols):
    for key in keywords:
        for idx, col in enumerate(cols):
            if key.lower() in str(col).lower().replace(" ", ""):
                return idx
    return 0

# 2. 파일 업로드 및 로드
uploaded_file = st.file_uploader("엑셀 또는 CSV 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

if uploaded_file is not None:
    if st.session_state.df_data is None:
        if uploaded_file.name.endswith('.csv'):
            try: st.session_state.df_data = pd.read_csv(uploaded_file, encoding='utf-8')
            except: st.session_state.df_data = pd.read_csv(uploaded_file, encoding='cp949')
        else:
            st.session_state.df_data = pd.read_excel(uploaded_file)
        
        st.session_state.df_data = st.session_state.df_data.loc[:, ~st.session_state.df_data.columns.duplicated()]
        if "입고예정수량(리오더)" not in st.session_state.df_data.columns:
            st.session_state.df_data["입고예정수량(리오더)"] = 0

if st.session_state.df_data is not None:
    df = st.session_state.df_data
    columns = df.columns.tolist()

    # [매핑 섹션]
    st.subheader("⚙️ 1단계: 자동 매핑 확인")
    col1, col2 = st.columns(2)
    with col1:
        sold_out = st.selectbox("품절 여부", columns, index=get_best_match(['품절', '판매중단'], columns))
        vendor = st.selectbox("공급처", columns, index=get_best_match(['공급처', '업체명'], columns))
        item = st.selectbox("상품명", columns, index=get_best_match(['상품명', '상품'], columns))
    with col2:
        avail = st.selectbox("가용재고", columns, index=get_best_match(['가용재고', '가용'], columns))
        t3day = st.selectbox("3일 발주 합계", columns, index=get_best_match(['3일', '최근3일'], columns))

    # [분석/필터링 섹션]
    st.write("---")
    st.subheader("⚙️ 2단계: 분석 및 필터링")
    c1, c2, c3 = st.columns([2, 1, 1])
    with c1: search_term = st.text_input("🔍 상품명 검색")
    with c2: filter_option = st.selectbox("품절 관리", ["전체 보기", "품절만 보기", "품절 제외"])
    with c3:
        if st.button("🚀 분석 실행"):
            df['일일 판매량(기준)'] = (pd.to_numeric(df[t3day], errors='coerce') / 3).round(0)
            st.session_state.df_data = df
            st.rerun()

    # 데이터 필터링 적용
    df_show = st.session_state.df_data.copy()
    if filter_option == "품절만 보기": df_show = df_show[df_show[sold_out].astype(str).str.upper() == 'Y']
    elif filter_option == "품절 제외": df_show = df_show[df_show[sold_out].astype(str).str.upper() != 'Y']
    if search_term: df_show = df_show[df_show[item].astype(str).str.contains(search_term, na=False)]

    # [결과 편집]
    st.subheader("📊 데이터 편집 및 결과 확인")
    
    edited_df = st.data_editor(df_show, use_container_width=True)
    
    # 수정된 데이터 원본 동기화
    st.session_state.df_data.update(edited_df)

    # [다운로드]
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        st.session_state.df_data.to_excel(writer, index=False)
    st.download_button("📥 최종 결과 엑셀 다운로드", data=buffer.getvalue(), file_name="최종_발주서.xlsx")
