import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="재고 관리 및 발주 시스템", layout="wide")
st.title("📦 재고 관리 및 발주 시스템")

if 'df_data' not in st.session_state:
    st.session_state.df_data = None

def get_best_match(keywords, cols):
    for key in keywords:
        for idx, col in enumerate(cols):
            if key.lower() in str(col).lower().replace(" ", ""):
                return idx
    return 0

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
            df["입고예정수량(리오더)"] = 0
            st.session_state.df_data = df
        except Exception as e:
            st.error(f"파일 로드 오류: {e}")

if st.session_state.df_data is not None:
    df = st.session_state.df_data
    columns = df.columns.tolist()

    st.subheader("⚙️ 1단계: 자동 매핑 확인")
    col1, col2 = st.columns(2)
    with col1:
        sold_out = st.selectbox("품절 여부", columns, index=get_best_match(['품절', '판매중단'], columns))
        vendor = st.selectbox("공급처", columns, index=get_best_match(['공급처', '업체명'], columns))
        item = st.selectbox("상품명", columns, index=get_best_match(['상품명', '상품'], columns))
        option = st.selectbox("옵션", columns, index=get_best_match(['옵션'], columns))
        vendor_option = st.selectbox("공급처옵션", columns, index=get_best_match(['공급처옵션', '거래처옵션'], columns))
    with col2:
        stock = st.selectbox("정상재고", columns, index=get_best_match(['정상재고', '재고'], columns))
        avail = st.selectbox("가용재고", columns, index=get_best_match(['가용재고', '가용'], columns))
        t3day = st.selectbox("3일 발주 합계", columns, index=get_best_match(['3일', '최근3일'], columns))
        t1week = st.selectbox("1주 발주 합계", columns, index=get_best_match(['1주', '7일', '최근7일'], columns))

    # [복구] 리드타임 및 안전재고 입력창
    st.write("---")
    st.subheader("⚙️ 2단계: 기간 설정")
    c1, c2 = st.columns(2)
    with c1: lead_time = st.number_input("평균 리드타임 (일)", min_value=0, value=0)
    with c2: safety_stock = st.number_input("안전재고 확보 기간 (일)", min_value=0, value=3)

    if st.button("🚀 분석 실행"):
        for col in [t3day, avail, "입고예정수량(리오더)"]:
            st.session_state.df_data[col] = pd.to_numeric(st.session_state.df_data[col], errors='coerce').fillna(0).astype(int)
            
        st.session_state.df_data['일일 판매량(기준)'] = (st.session_state.df_data[t3day] / 3).round(0).astype(int)
        
        # 권장 발주량 계산 공식: (일일판매량 * (리드타임 + 안전재고)) - (가용재고 + 리오더)
        st.session_state.df_data['권장 발주량'] = (st.session_state.df_data['일일 판매량(기준)'] * (lead_time + safety_stock) - 
                                             (st.session_state.df_data[avail] + st.session_state.df_data["입고예정수량(리오더)"])).clip(lower=0).astype(int)
        st.rerun()

    st.subheader("📊 데이터 편집 및 결과 확인")
    
    
    display_cols = [sold_out, vendor, item, option, vendor_option, stock, avail, "입고예정수량(리오더)", t3day, t1week, '일일 판매량(기준)', '권장 발주량']
    result_df = st.session_state.df_data[[c for c in display_cols if c in st.session_state.df_data.columns]]
    
    edited_df = st.data_editor(result_df, column_config={
        "입고예정수량(리오더)": st.column_config.NumberColumn(format="%d")
    }, use_container_width=True)
    
    for col in edited_df.columns:
        st.session_state.df_data[col] = edited_df[col]

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        st.session_state.df_data.to_excel(writer, index=False)
    
    st.download_button("📥 최종 결과 엑셀 다운로드", data=buffer.getvalue(), file_name="최종_발주서.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
