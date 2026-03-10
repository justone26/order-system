import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="재고 관리 대시보드", layout="wide")
st.title("📦 3일 발주 합계 기반 정밀 재고 관리 시스템")

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
        vendor_option = st.selectbox("공급처옵션", columns, index=get_default_index(['공급처옵션'], columns))
    with col2:
        stock_col = st.selectbox("정상재고", columns, index=get_default_index(['정상'], columns))
        avail_col = st.selectbox("가용재고", columns, index=get_default_index(['가용'], columns))
        target_3day = st.selectbox("3일 발주 합계", columns, index=get_default_index(['3일'], columns))
        lead_time_input = st.number_input("평균 리드타임 (일)", min_value=0, value=7)
        safety_stock_input = st.number_input("안전재고 확보일 (일)", min_value=0, value=3)

    if st.button("분석 및 이력 저장"):
        try:
            for col in [stock_col, avail_col, target_3day]:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

            # 운영 설정값 대입
            df['평균 리드타임(일)'] = lead_time_input
            df['안전재고 확보일(일)'] = safety_stock_input
            df['총 확보 기간(일)'] = lead_time_input + safety_stock_input

            # 데이터 산출
            df['일일 판매량(기준)'] = (df[target_3day] / 3).round(0).astype(int)
            df['일일 판매량(재고차액)'] = (df[stock_col] - df[avail_col]).clip(lower=0).astype(int)
            
            # 발주 수량 계산
            df['1일 추가 발주 필요수량'] = ((df['일일 판매량(기준)'] * 1) - df[avail_col]).clip(lower=0).astype(int)
            df['3일 추가 발주 필요수량'] = ((df['일일 판매량(기준)'] * 3) - df[avail_col]).clip(lower=0).astype(int)
            df['7일 추가 발주 필요수량'] = ((df['일일 판매량(기준)'] * 7) - df[avail_col]).clip(lower=0).astype(int)
            df['권장 발주수량(리드타임)'] = ((df['일일 판매량(기준)'] * df['총 확보 기간(일)']) - df[avail_col]).clip(lower=0).astype(int)

            df['상태'] = df.apply(lambda row: '🚨 품절/긴급' if (str(row[sold_out_col]).upper() == 'Y' or row[avail_col] <= 0) else '정상', axis=1)
            
            df['저장날짜'] = pd.Timestamp.now().strftime('%Y-%m-%d')
            df.to_csv('order_history.csv', mode='a', header=not os.path.exists('order_history.csv'), index=False)
            
            st.session_state.analysis_result = df
            st.success("데이터 분석 및 설정값 반영 완료!")
        except Exception as e:
            st.error(f"오류 발생: {e}")

    if st.session_state.analysis_result is not None:
        st.subheader("📊 분석 결과 대시보드")
        
        def highlight_status(df):
            styles = pd.DataFrame('', index=df.index, columns=df.columns)
            if '상태' in df.columns:
                mask = df['상태'] == '🚨 품절/긴급'
                styles.loc[mask, :] = 'background-color: #ffcccc'
            return styles
        
        st.dataframe(st.session_state.analysis_result.style.apply(highlight_status, axis=None))
