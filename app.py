import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="자동 발주 및 재고 관리", layout="wide")
st.title("📦 리드타임/안전재고 반영 발주 관리 시스템")

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
    st.subheader("⚙️ 데이터 매핑 및 운영 설정")
    col1, col2 = st.columns(2)
    with col1:
        sold_out_col = st.selectbox("품절 여부", columns, index=get_default_index(['품절'], columns))
        vendor_name = st.selectbox("공급처", columns, index=get_default_index(['공급처'], columns))
        item_name = st.selectbox("상품명", columns, index=get_default_index(['상품'], columns))
        option_name = st.selectbox("옵션", columns, index=get_default_index(['옵션'], columns))
        stock_col = st.selectbox("정상재고", columns, index=get_default_index(['재고'], columns))
    with col2:
        col_invoice = st.selectbox("송장", columns, index=get_default_index(['송장'], columns))
        col_reception = st.selectbox("접수", columns, index=get_default_index(['접수'], columns))
        target_3day = st.selectbox("3일 발주 합계", columns, index=get_default_index(['3일'], columns))
        lead_time = st.number_input("평균 리드타임 (일)", min_value=0, value=7)
        safety_stock_days = st.number_input("안전재고 확보일 (일)", min_value=0, value=3)

    if st.button("분석 및 이력 저장"):
        try:
            for col in [col_invoice, col_reception, target_3day, stock_col]:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            df['일 판매 데이터'] = df[col_invoice] + df[col_reception]
            df['일일평균'] = (df['일 판매 데이터'] / 1).round(1)
            df['재고소진일'] = (df[stock_col] / df['일일평균'].replace(0, 1)).round(1)
            
            df['3일 필요수량'] = ((df['일일평균'] * 3) - df[stock_col]).clip(lower=0)
            df['권장발주수량'] = ((df['일일평균'] * (lead_time + safety_stock_days)) - df[stock_col]).clip(lower=0)

            df['상태'] = df.apply(lambda row: '🚨 품절/긴급' if (str(row[sold_out_col]).upper() == 'Y' or row[stock_col] <= 0 or row['재고소진일'] <= 2) else '정상', axis=1)
            
            df['저장날짜'] = pd.Timestamp.now().strftime('%Y-%m-%d')
            df.to_csv('order_history.csv', mode='a', header=not os.path.exists('order_history.csv'), index=False)
            
            st.session_state.analysis_result = df
            st.success("분석 완료!")
        except Exception as e:
            st.error(f"오류: {e}")

    if st.session_state.analysis_result is not None:
        st.subheader("📊 최신 분석 결과")
        
        def highlight_status(row):
            # '상태' 컬럼이 존재하는지 명확히 확인 후 색상 적용
            if '상태' in row.index and row['상태'] == '🚨 품절/긴급':
                return ['background-color: #ffcccc'] * len(row)
            return [''] * len(row)
        
        st.dataframe(st.session_state.analysis_result.style.apply(highlight_status, axis=1))

    st.write("---")
    st.subheader("📅 과거 발주 내역 검색")
    search_date = st.date_input("조회할 날짜 선택")
    if st.button("내역 조회"):
        if os.path.exists('order_history.csv'):
            history_df = pd.read_csv('order_history.csv')
            filtered_df = history_df[history_df['저장날짜'] == str(search_date)]
            if not filtered_df.empty:
                st.dataframe(filtered_df)
            else:
                st.warning("데이터가 없습니다.")
