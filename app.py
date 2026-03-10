import streamlit as st
import pandas as pd

st.set_page_config(page_title="재고 관리 대시보드", layout="wide")
st.title("📦 재고 관리 및 발주 시스템")

uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요", type=['xlsx', 'xls'])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    columns = df.columns.tolist()

    def get_default_index(keywords, cols):
        for col in cols:
            for key in keywords:
                if key in col: return cols.index(col)
        return 0

    st.write("---")
    st.subheader("⚙️ 1단계: 데이터 매핑")
    col1, col2 = st.columns(2)
    with col1:
        item_name = st.selectbox("상품명", columns, index=get_default_index(['상품'], columns))
        option_name = st.selectbox("옵션", columns, index=get_default_index(['옵션'], columns))
        vendor = st.selectbox("공급처", columns, index=get_default_index(['공급처'], columns))
    with col2:
        stock_col = st.selectbox("정상재고", columns, index=get_default_index(['정상'], columns))
        avail_col = st.selectbox("가용재고", columns, index=get_default_index(['가용'], columns))
        target_3day = st.selectbox("3일 발주 합계", columns, index=get_default_index(['3일'], columns))

    st.write("---")
    st.subheader("⚙️ 2단계: 기간 기반 산출 설정")
    col3, col4 = st.columns(2)
    with col3:
        lead_time_days = st.selectbox("평균 리드타임 기간(일)", options=[3, 5, 7, 10, 14, 21, 30], index=2)
    with col4:
        safety_stock_days = st.selectbox("안전재고 확보 기간(일)", options=[1, 2, 3, 5, 7, 10], index=2)

    if st.button("분석 실행"):
        # 1. 일일 판매량(기준): 3일 발주 합계 기반
        df['일일 판매량(기준)'] = (df[target_3day] / 3).round(0).astype(int)
        
        # 2. 필수 재고 수량 계산
        # 리드타임 준비량 = 일일판매량 * 리드타임 기간
        # 안전재고 수량 = 일일판매량 * 안전재고 기간
        df['리드타임 준비량'] = (df['일일 판매량(기준)'] * lead_time_days).astype(int)
        df['안전재고 수량'] = (df['일일 판매량(기준)'] * safety_stock_days).astype(int)
        
        # 3. 최종 권장 발주량: (준비량 + 안전재고) - 가용재고
        df['권장 발주량'] = (df['리드타임 준비량'] + df['안전재고 수량'] - df[avail_col]).clip(lower=0).astype(int)
        
        st.subheader("📊 분석 결과")
        
        st.dataframe(df)
