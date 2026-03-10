import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="이동평균 기반 발주 시스템", layout="wide")
st.title("📦 3일/7일 이동평균 기반 발주 관리 시스템")

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
    with col2:
        avail_col = st.selectbox("가용재고", columns, index=get_default_index(['가용'], columns))
        target_3day = st.selectbox("3일 발주 합계", columns, index=get_default_index(['3일'], columns))
        target_7day = st.selectbox("1주 발주 합계", columns, index=get_default_index(['1주', '7일'], columns))
        lead_time = st.number_input("평균 리드타임 (일)", min_value=0, value=7)
        safety_stock_days = st.number_input("안전재고 확보일 (일)", min_value=0, value=3)

    if st.button("분석 및 이력 저장"):
        try:
            for col in [target_3day, target_7day, avail_col]:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            # 이동평균 기반 판매량 산출
            df['판매량_기준'] = (((df[target_3day] / 3) + (df[target_7day] / 7)) / 2).round(1)
            
            # 발주 필요수량 계산
            df['3일 추가 발주 필요수량'] = ((df[target_3day] / 3 * 3) - df[avail_col]).clip(lower=0).astype(int)
            df['7일 추가 발주 필요수량'] = ((df[target_7day] / 7 * 7) - df[avail_col]).clip(lower=0).astype(int)
            df['권장 발주수량(리드타임)'] = ((df['판매량_기준'] * (lead_time + safety_stock_days)) - df[avail_col]).clip(lower=0).astype(int)

            df['저장날짜'] = pd.Timestamp.now().strftime('%Y-%m-%d')
            df.to_csv('order_history.csv', mode='a', header=not os.path.exists('order_history.csv'), index=False)
            
            st.session_state.analysis_result = df
            st.success("이동평균 기반 분석 완료!")
        except Exception as e:
            st.error(f"오류 발생: {e}")

    if st.session_state.analysis_result is not None:
        st.subheader("📊 분석 결과")
        st.dataframe(st.session_state.analysis_result)
