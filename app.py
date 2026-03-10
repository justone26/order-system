import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="기간별 필수재고 산출 시스템", layout="wide")
st.title("📦 기간별 필수재고 산출 및 발주 관리")

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
    st.subheader("⚙️ 분석 설정")
    
    col1, col2 = st.columns(2)
    with col1:
        stock_col = st.selectbox("정상재고", columns, index=get_default_index(['정상'], columns))
        avail_col = st.selectbox("가용재고", columns, index=get_default_index(['가용'], columns))
        target_3day = st.selectbox("3일 발주 합계", columns, index=get_default_index(['3일'], columns))
    with col2:
        # [핵심] 기간 선택 셀렉터
        period_days = st.selectbox("필수 재고 산출 기간(일)", options=[1, 3, 7, 10, 14, 30], index=1)
        lead_time = st.number_input("평균 리드타임 (일)", min_value=0, value=7)
        safety_stock_qty = st.number_input("고정 안전재고 수량 (개)", min_value=0, value=10)

    if st.button("분석 실행"):
        try:
            for col in [stock_col, avail_col, target_3day]:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

            # 1. 일일 판매량 산출 (3일 합계 기준)
            df['일일 판매량'] = (df[target_3day] / 3).round(0).astype(int)
            
            # 2. [요청 로직] 선택 기간별 필수 재고량 산출
            # 필수 재고량 = (일일판매량 * 선택기간) + 안전재고
            df[f'{period_days}일 필수재고량'] = (df['일일 판매량'] * period_days) + safety_stock_qty
            
            # 3. 발주 필요 수량 (필수 재고 - 가용재고)
            df['추가 발주 필요수량'] = (df[f'{period_days}일 필수재고량'] - df[avail_col]).clip(lower=0)

            st.session_state.analysis_result = df
            st.success(f"{period_days}일 기준 필수 재고량 산출 완료!")
        except Exception as e:
            st.error(f"오류 발생: {e}")

    if st.session_state.analysis_result is not None:
        st.subheader("📊 분석 결과")
        
        
        
        st.dataframe(st.session_state.analysis_result)
