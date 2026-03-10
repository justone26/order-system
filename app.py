import streamlit as st
import pandas as pd
import os
import io

st.set_page_config(page_title="자동 발주 및 재고 관리", layout="wide")
st.title("📦 재고 기반 판매량 분석 시스템")

if 'analysis_result' not in st.session_state: st.session_state.analysis_result = None

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
    st.subheader("⚙️ 데이터 항목 매핑")
    col1, col2 = st.columns(2)
    with col1:
        item_name = st.selectbox("상품명", columns, index=get_default_index(['상품', '품명'], columns))
        option_name = st.selectbox("옵션", columns, index=get_default_index(['옵션'], columns))
        # 재고 기반 계산을 위한 필수 항목
        yesterday_stock = st.selectbox("어제 가용재고", columns, index=get_default_index(['어제', '전일'], columns))
        today_stock = st.selectbox("오늘 가용재고", columns, index=get_default_index(['오늘', '당일'], columns))
        stock_col = st.selectbox("현재 정상재고", columns, index=get_default_index(['재고', '정상'], columns))
    with col2:
        target_3day = st.selectbox("3일 발주 합계", columns, index=get_default_index(['3일'], columns))
        # 기타 필요한 매핑...

    if st.button("분석 및 이력 저장"):
        try:
            # [핵심 로직 변경] 재고 차액으로 판매량 산출
            df['일 판매 데이터'] = df[yesterday_stock] - df[today_stock]
            df['일일평균'] = df[target_3day] / 3
            df['재고소진일'] = (df[stock_col] / df['일일평균'].replace(0, 1)).round(1)
            
            df['저장날짜'] = pd.Timestamp.now().strftime('%Y-%m-%d')
            df.to_csv('order_history.csv', mode='a', header=not os.path.exists('order_history.csv'), index=False)
            
            st.session_state.analysis_result = df
            st.success("재고 차액 기반 분석 완료!")
        except Exception as e:
            st.error(f"오류 발생: {e}")

    if st.session_state.analysis_result is not None:
        st.subheader("📊 최신 분석 결과")
        st.dataframe(st.session_state.analysis_result)
        # ... (다운로드 버튼 및 추이 분석 로직 동일)
