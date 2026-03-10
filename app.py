import streamlit as st
import pandas as pd
import os
import io

st.set_page_config(page_title="자동 발주 및 재고 관리", layout="wide")
st.title("📦 리드타임/안전재고 반영 발주 관리 시스템")

# [세션 상태 관리]
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
        vendor_name = st.selectbox("공급처", columns, index=get_default_index(['공급처'], columns))
        item_name = st.selectbox("상품명", columns, index=get_default_index(['상품'], columns))
        option_name = st.selectbox("옵션", columns, index=get_default_index(['옵션'], columns))
        stock_col = st.selectbox("정상재고", columns, index=get_default_index(['재고'], columns))
        available_stock = st.selectbox("가용재고", columns, index=get_default_index(['가용'], columns))
    with col2:
        col_invoice = st.selectbox("송장", columns, index=get_default_index(['송장'], columns))
        col_reception = st.selectbox("접수", columns, index=get_default_index(['접수'], columns))
        target_3day = st.selectbox("3일 발주 합계", columns, index=get_default_index(['3일'], columns))
        
        lead_time = st.number_input("평균 리드타임 (일)", min_value=0, value=7)
        safety_stock_days = st.number_input("안전재고 확보일 (일)", min_value=0, value=3)

    if st.button("분석 및 발주량 계산"):
        try:
            # 수치 데이터 정제
            for col in [col_invoice, col_reception, target_3day, stock_col, available_stock]:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            # 일 판매 데이터 계산
            df['일 판매 데이터'] = df[col_invoice] + df[col_reception]
            df['일일평균'] = (df[target_3day] / 3).round(1)
            
            # 리드타임 + 안전재고 반영 발주 공식
            # 공식: (일일평균 * (리드타임 + 안전재고)) - 현재가용재고
            df['권장발주수량'] = (df['일일평균'] * (lead_time + safety_stock_days) - df[available_stock]).apply(lambda x: max(0, int(x)))
            
            df['저장날짜'] = pd.Timestamp.now().strftime('%Y-%m-%d')
            st.session_state.analysis_result = df
            st.success(f"분석 완료! (리드타임 {lead_time}일 + 안전재고 {safety_stock_days}일 기준)")
        except Exception as e:
            st.error(f"오류 발생: {e}")

    # 데이터 시각화 및 출력 섹션
    if st.session_state.analysis_result is not None:
        st.subheader("📊 발주 분석 대시보드")
        
        # 
        
        # 권장 발주량 시각화 (데이터 존재 여부 확인 후 실행)
        if '권장발주수량' in st.session_state.analysis_result.columns:
            st.bar_chart(st.session_state.analysis_result.set_index(option_name)['권장발주수량'])
            
        st.dataframe(st.session_state.analysis_result)
        
        # 엑셀 다운로드
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            st.session_state.analysis_result.to_excel(writer, index=False)
        st.download_button("📥 결과 파일 다운로드 (Excel)", buffer.getvalue(), "발주계산결과.xlsx")
