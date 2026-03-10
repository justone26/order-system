import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="재고 관리 대시보드", layout="wide")
st.title("📦 재고 최적화 및 발주 관리 시스템")

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
        item_name = st.selectbox("상품명", columns, index=get_default_index(['상품'], columns))
        stock_col = st.selectbox("정상재고", columns, index=get_default_index(['정상'], columns))
        avail_col = st.selectbox("가용재고", columns, index=get_default_index(['가용'], columns))
    with col2:
        target_3day = st.selectbox("3일 발주 합계", columns, index=get_default_index(['3일'], columns))
        lead_time = st.number_input("평균 리드타임 (일)", min_value=0, value=7)
        # [변경] 일수가 아닌 '수량'을 직접 설정하도록 변경
        safety_stock_qty = st.number_input("안전재고 필요 수량 (개수)", min_value=0, value=10)

    if st.button("분석 및 이력 저장"):
        try:
            for col in [stock_col, avail_col, target_3day]:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

            # 일일 판매량(기준)
            df['일일 판매량(기준)'] = (df[target_3day] / 3).round(0).astype(int)
            
            # [핵심 로직] 권장 발주량 = (입고 대기기간 판매분) + 안전재고 수량 - 현재 가용재고
            # 입고 대기기간 판매분 = 일일판매량 * 리드타임
            lead_time_sales = df['일일 판매량(기준)'] * lead_time
            df['권장 발주수량(리드타임)'] = (lead_time_sales + safety_stock_qty - df[avail_col]).clip(lower=0).astype(int)
            
            # 발주 필요수량 계산 (1일/3일/7일 기준은 판매량 기반 그대로 유지)
            df['1일 추가 발주 필요수량'] = ((df['일일 판매량(기준)'] * 1) - df[avail_col]).clip(lower=0).astype(int)
            df['3일 추가 발주 필요수량'] = ((df['일일 판매량(기준)'] * 3) - df[avail_col]).clip(lower=0).astype(int)
            df['7일 추가 발주 필요수량'] = ((df['일일 판매량(기준)'] * 7) - df[avail_col]).clip(lower=0).astype(int)
            
            # 운영 정보 기록용
            df['설정 안전재고(개수)'] = safety_stock_qty
            
            # 상태 판정
            df['상태'] = df.apply(lambda row: '🚨 품절/긴급' if (str(row[sold_out_col]).upper() == 'Y' or row[avail_col] <= 0) else '정상', axis=1)
            
            st.session_state.analysis_result = df
            st.success("데이터 분석 완료!")
        except Exception as e:
            st.error(f"오류 발생: {e}")

    if st.session_state.analysis_result is not None:
        st.subheader("📊 분석 결과 대시보드")
        
        # 
        
        st.dataframe(st.session_state.analysis_result)
