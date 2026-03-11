import streamlit as st
import pandas as pd
from io import BytesIO

# 페이지 설정
st.set_page_config(page_title="재고 관리 및 발주 시스템", layout="wide")
st.title("📦 재고 관리 및 발주 시스템")

# [세션 상태 관리: 데이터 유지]
if 'df_data' not in st.session_state:
    st.session_state.df_data = None

# 자동 매핑 함수
def get_best_match(keywords, cols):
    for key in keywords:
        for idx, col in enumerate(cols):
            if key.lower() in str(col).lower().replace(" ", ""):
                return idx
    return 0

# 1. 파일 업로드 및 데이터 로드
uploaded_file = st.file_uploader("엑셀 또는 CSV 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

if uploaded_file is not None:
    if st.session_state.df_data is None:
        try:
            if uploaded_file.name.endswith('.csv'):
                try: df = pd.read_csv(uploaded_file, encoding='utf-8')
                except: df = pd.read_csv(uploaded_file, encoding='cp949')
            else:
                df = pd.read_excel(uploaded_file)
            
            df = df.loc[:, ~df.columns.duplicated()]
            if "입고예정수량(리오더)" not in df.columns:
                df["입고예정수량(리오더)"] = 0
            
            st.session_state.df_data = df
        except Exception as e:
            st.error(f"파일 로드 오류: {e}")

if st.session_state.df_data is not None:
    df = st.session_state.df_data
    columns = df.columns.tolist()

    # 2. 자동 매핑 설정 (항상 유지)
    st.subheader("⚙️ 1단계: 자동 매핑 확인")
    col1, col2 = st.columns(2)
    with col1:
        sold_out = st.selectbox("품절 여부", columns, index=get_best_match(['품절', '판매중단'], columns))
        vendor = st.selectbox("공급처", columns, index=get_best_match(['공급처', '업체명'], columns))
        item = st.selectbox("상품명", columns, index=get_best_match(['상품명', '상품'], columns))
        option = st.selectbox("옵션", columns, index=get_best_match(['옵션'], columns))
    with col2:
        stock = st.selectbox("정상재고", columns, index=get_best_match(['정상재고', '재고'], columns))
        avail = st.selectbox("가용재고", columns, index=get_best_match(['가용재고', '가용'], columns))
        t3day = st.selectbox("3일 발주 합계", columns, index=get_best_match(['3일', '최근3일'], columns))
        t1week = st.selectbox("1주 발주 합계", columns, index=get_best_match(['1주', '7일', '최근7일'], columns))

    # 3. 분석 실행
    if st.button("🚀 분석 실행"):
        for col in [t3day, avail, "입고예정수량(리오더)"]:
            st.session_state.df_data[col] = pd.to_numeric(st.session_state.df_data[col], errors='coerce').fillna(0)
            
        st.session_state.df_data['일일 판매량(기준)'] = (st.session_state.df_data[t3day] / 3).round(0).astype(int)
        st.session_state.df_data['권장 발주량'] = (st.session_state.df_data['일일 판매량(기준)'] * 3 - 
                                             (st.session_state.df_data[avail] + st.session_state.df_data["입고예정수량(리오더)"])).clip(lower=0).astype(int)
        st.rerun()

    # 4. 결과 출력 및 수동 편집 (필수 항목만 필터링)
    st.subheader("📊 데이터 편집 및 결과 확인")
    
    # 핵심 데이터만 추림
    display_cols = [sold_out, vendor, item, option, stock, avail, "입고예정수량(리오더)", t3day, t1week, '일일 판매량(기준)', '권장 발주량']
    result_df = st.session_state.df_data[[c for c in display_cols if c in st.session_state.df_data.columns]]

    
    
    # 데이터 편집기
    edited_df = st.data_editor(result_df, use_container_width=True)
    
    # 편집된 값을 원본 데이터와 동기화
    for col in edited_df.columns:
        st.session_state.df_data[col] = edited_df[col]

    # 5. 엑셀 다운로드
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        st.session_state.df_data.to_excel(writer, index=False)
    
    st.download_button("📥 수정된 결과 엑셀 다운로드", data=buffer.getvalue(), file_name="최종_발주서.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
