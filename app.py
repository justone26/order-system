import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="자동 발주 시스템", layout="wide")
st.title("📦 재고 품절 방지 알림 시스템 (2일/4일/7일 발주)")

uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요", type=['xlsx', 'xls'])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    columns = df.columns.tolist()

    st.write("---")
    st.subheader("⚙️ 데이터 항목 매핑")
    
    col1, col2 = st.columns(2)
    with col1:
        col_invoice = st.selectbox("송장 항목 선택", columns)
        col_reception = st.selectbox("접수 항목 선택", columns)
        stock_col = st.selectbox("현재고(정상재고) 항목 선택", columns)
    with col2:
        vendor_col = st.selectbox("거래처명 항목 선택", columns)
        target_3day = st.selectbox("3일 발주 합계 항목 선택", columns)

    if st.button("분석 및 경고 확인"):
        try:
            # 1. 1일 판매량 합산 및 일일 평균 계산
            df['1일판매량'] = df[col_invoice] + df[col_reception]
            df['일일평균'] = df[target_3day] / 3
            
            # 2. 재고 소진일 계산
            df['재고소진일'] = (df[stock_col] / df['일일평균'].replace(0, 1)).round(1)
            
            # 3. 발주량 계산 (2일분, 4일분, 7일분)
            df['2일권장발주'] = (df['일일평균'] * 2 - df[stock_col]).apply(lambda x: max(0, int(x)))
            df['4일권장발주'] = (df['일일평균'] * 4 - df[stock_col]).apply(lambda x: max(0, int(x)))
            df['7일권장발주'] = (df['일일평균'] * 7 - df[stock_col]).apply(lambda x: max(0, int(x)))
            
            # 4. 상태 표기 (3일 미만 긴급)
            df['상태'] = df['재고소진일'].apply(lambda x: '🚨 긴급(3일미만)' if x < 3 else '정상')
            
            st.success("데이터 분석 완료!")
            
            # 5. 결과 표시 (2일/4일/7일 발주량 포함)
            cols_to_show = ['상품명', '옵션', '현재고', '일일평균', '2일권장발주', '4일권장발주', '7일권장발주', '상태']
            # 데이터프레임에서 선택한 컬럼이 없는 경우 전체 표시
            st.dataframe(df)

            # 6. 엑셀 다운로드 (openpyxl 사용)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 결과 파일 다운로드 (Excel)", 
                data=output.getvalue(), 
                file_name="분석결과_발주서.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"계산 중 오류 발생: {e}")