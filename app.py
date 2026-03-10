import streamlit as st
import pandas as pd

st.set_page_config(page_title="재고 관리 및 판매 분석", layout="wide")
st.title("📦 지능형 재고 관리 시스템")

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
        sold_out_col = st.selectbox("품절 여부", columns, index=get_default_index(['품절'], columns))
        vendor_col = st.selectbox("공급처", columns, index=get_default_index(['공급처'], columns))
        item_name = st.selectbox("상품명", columns, index=get_default_index(['상품'], columns))
        option_name = st.selectbox("옵션", columns, index=get_default_index(['옵션'], columns))
        vendor_option = st.selectbox("공급처옵션", columns, index=get_default_index(['공급처옵션'], columns))
    with col2:
        stock_col = st.selectbox("정상재고", columns, index=get_default_index(['정상'], columns))
        avail_col = st.selectbox("가용재고", columns, index=get_default_index(['가용'], columns))
        target_3day = st.selectbox("3일 발주 합계", columns, index=get_default_index(['3일'], columns))
        target_1week = st.selectbox("1주 발주 합계", columns, index=get_default_index(['1주', '7일'], columns))

    st.write("---")
    st.subheader("⚙️ 2단계: 기간 기반 산출 설정")
    col3, col4 = st.columns(2)
    with col3:
        lead_time_days = st.number_input("평균 리드타임 기간 (일)", min_value=0, value=7)
    with col4:
        safety_stock_days = st.number_input("안전재고 확보 기간 (일)", min_value=0, value=3)

    if st.button("분석 실행"):
        try:
            # 1. 수치 데이터 정제
            for col in [stock_col, avail_col, target_3day, target_1week]:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            # 2. 일일 판매량(기준) 산출
            df['일일 판매량(기준)'] = (df[target_3day] / 3).round(1)
            
            # 3. 리드타임 및 안전재고 수량 산출
            df['리드타임 준비량'] = (df['일일 판매량(기준)'] * lead_time_days).astype(int)
            df['안전재고 수량'] = (df['일일 판매량(기준)'] * safety_stock_days).astype(int)
            
            # 4. 권장 발주량: (준비량 + 안전재고) - 가용재고
            df['권장 발주량'] = (df['리드타임 준비량'] + df['안전재고 수량'] - df[avail_col]).clip(lower=0).astype(int)
            
            # 5. 과거 데이터 대조 (판매 추이 분석: 1주 평균 대비 최근 3일 속도)
            df['판매 추이(1=평균)'] = (df['일일 판매량(기준)'] / (df[target_1week] / 7 + 0.01)).round(2)

            st.subheader("📊 분석 결과 리스트")
            # 주요 컬럼을 앞쪽으로 배치하여 가독성 향상
            display_cols = [vendor_col, item_name, option_name, '일일 판매량(기준)', '리드타임 준비량', '안전재고 수량', avail_col, '권장 발주량', '판매 추이(1=평균)']
            st.dataframe(df[display_cols])

        except Exception as e:
            st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")
