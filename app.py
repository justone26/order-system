import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="재고 관리 대시보드", layout="wide")
st.title("📦 이동평균 및 재고 차액 분석 시스템")

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
        stock_col = st.selectbox("정상재고", columns, index=get_default_index(['정상'], columns))
        avail_col = st.selectbox("가용재고", columns, index=get_default_index(['가용'], columns))
        target_3day = st.selectbox("3일 발주 합계", columns, index=get_default_index(['3일'], columns))
        target_7day = st.selectbox("1주 발주 합계", columns, index=get_default_index(['1주', '7일'], columns))
    with col2:
        lead_time = st.number_input("평균 리드타임 (일)", min_value=0, value=7)
        safety_stock_days = st.number_input("안전재고 확보일 (일)", min_value=0, value=3)

    if st.button("분석 및 이력 저장"):
        try:
            for col in [stock_col, avail_col, target_3day, target_7day]:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            # 1. 일일 판매량 (정상 - 가용)
            df['일일 판매량(재고차액)'] = (df[stock_col] - df[avail_col]).clip(lower=0)
            
            # 2. 이동평균 기반 판매량 산출
            df['이동평균_판매량'] = (((df[target_3day] / 3) + (df[target_7day] / 7)) / 2).round(1)
            
            # 3. 발주 필요수량 계산 (이동평균 기준)
            df['3일 추가 발주 필요수량'] = ((df['이동평균_판매량'] * 3) - df[avail_col]).clip(lower=0).astype(int)
            df['7일 추가 발주 필요수량'] = ((df['이동평균_판매량'] * 7) - df[avail_col]).clip(lower=0).astype(int)
            df['권장 발주수량(리드타임)'] = ((df['이동평균_판매량'] * (lead_time + safety_stock_days)) - df[avail_col]).clip(lower=0).astype(int)

            df['저장날짜'] = pd.Timestamp.now().strftime('%Y-%m-%d')
            df.to_csv('order_history.csv', mode='a', header=not os.path.exists('order_history.csv'), index=False)
            
            st.session_state.analysis_result = df
            st.success("데이터 분석 완료!")
        except Exception as e:
            st.error(f"오류 발생: {e}")

    if st.session_state.analysis_result is not None:
        st.subheader("📊 분석 결과 대시보드")
        
        
        
        st.dataframe(st.session_state.analysis_result)
