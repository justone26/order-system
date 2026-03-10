import streamlit as st
import pandas as pd

st.set_page_config(page_title="기간 기반 재고 관리", layout="wide")
st.title("📦 리드타임/안전재고 기간 설정형 시스템")

uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요", type=['xlsx', 'xls'])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    columns = df.columns.tolist()

    # 데이터 매핑
    col1, col2 = st.columns(2)
    with col1:
        stock_col = st.selectbox("정상재고", columns, index=0)
        avail_col = st.selectbox("가용재고", columns, index=1)
    with col2:
        target_3day = st.selectbox("3일 발주 합계", columns, index=2)
        
    st.write("---")
    st.subheader("⚙️ 리드타임 및 안전재고 기간 설정")
    
    # [핵심] 기간을 선택할 수 있는 셀렉터/입력창
    col3, col4 = st.columns(2)
    with col3:
        # 리드타임 기간 선택 (예: 3일, 5일, 7일, 10일...)
        lead_time_days = st.selectbox("평균 리드타임 기간 선택(일)", options=[3, 5, 7, 10, 14, 21, 30], index=2)
    with col4:
        # 안전재고 기간 선택 (며칠 분의 물량을 확보할 것인가)
        safety_stock_days = st.selectbox("안전재고 확보 기간 선택(일)", options=[1, 2, 3, 5, 7, 10], index=2)

    if st.button("분석 실행"):
        # 1. 일일 판매량(기준): 3일 발주 합계 기준
        df['일일 판매량(기준)'] = (df[target_3day] / 3).round(0).astype(int)
        
        # 2. 리드타임 준비량: 선택한 리드타임 기간만큼의 판매량
        df['리드타임 준비량'] = (df['일일 판매량(기준)'] * lead_time_days).astype(int)
        
        # 3. 안전재고 수량: 선택한 안전재고 기간만큼의 판매량
        df['안전재고 수량'] = (df['일일 판매량(기준)'] * safety_stock_days).astype(int)
        
        # 4. 최종 권장 발주량
        df['권장 발주량'] = (df['리드타임 준비량'] + df['안전재고 수량'] - df[avail_col]).clip(lower=0).astype(int)
        
        # 결과 출력 (매핑값 확인용)
        st.subheader("📊 분석 결과")
        st.info(f"설정 적용: 리드타임 {lead_time_days}일 + 안전재고 {safety_stock_days}일분 반영")
        st.dataframe(df[['일일 판매량(기준)', '리드타임 준비량', '안전재고 수량', '권장 발주량']])
