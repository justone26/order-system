import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="자동 발주 시스템", layout="wide")
st.title("📦 재고 품절 방지 알림 시스템")

uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요", type=['xlsx', 'xls'])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    columns = df.columns.tolist()

    # 자동 매핑 도우미 함수
    def get_default_index(keywords, cols):
        for col in cols:
            for key in keywords:
                if key in col: return cols.index(col)
        return 0

    st.write("---")
    st.subheader("⚙️ 데이터 항목 매핑")
    
    col1, col2 = st.columns(2)
    with col1:
        # '품절' 항목 추가
        sold_out_col = st.selectbox("품절 여부", columns, index=get_default_index(['품절'], columns))
        item_name = st.selectbox("상품명", columns, index=get_default_index(['상품', '품명'], columns))
        option_name = st.selectbox("옵션", columns, index=get_default_index(['옵션'], columns))
        vendor_name = st.selectbox("공급처명", columns, index=get_default_index(['공급처', '거래처'], columns))
        vendor_option = st.selectbox("공급처옵션", columns, index=get_default_index(['공급처옵션'], columns))
        stock_col = st.selectbox("정상재고", columns, index=get_default_index(['재고', '정상'], columns))
    with col2:
        col_invoice = st.selectbox("송장", columns, index=get_default_index(['송장'], columns))
        col_reception = st.selectbox("접수", columns, index=get_default_index(['접수'], columns))
        target_3day = st.selectbox("3일 발주 합계", columns, index=get_default_index(['3일'], columns))
        target_1week = st.selectbox("1주 발주 합계", columns, index=get_default_index(['1주', '7일'], columns))

    if st.button("분석 및 경고 확인"):
        try:
            # 일 판매 데이터 합산
            df['일 판매 데이터'] = df[col_invoice] + df[col_reception]
            df['일일평균'] = df[target_3day] / 3
            df['재고소진일'] = (df[stock_col] / df['일일평균'].replace(0, 1)).round(1)
            
            # 발주량 계산
            df['2일권장발주'] = (df['일일평균'] * 2 - df[stock_col]).apply(lambda x: max(0, int(x)))
            df['4일권장발주'] = (df['일일평균'] * 4 - df[stock_col]).apply(lambda x: max(0, int(x)))
            df['7일권장발주'] = (df['일일평균'] * 7 - df[stock_col]).apply(lambda x: max(0, int(x)))
            
            # 상태 표기 (품절 체크 추가: 품절 항목이 'Y'이거나 재고가 0이면 긴급)
            df['상태'] = df.apply(lambda row: '🚨 품절/긴급' if (str(row[sold_out_col]).upper() == 'Y' or row[stock_col] <= 0 or row['재고소진일'] < 3) else '정상', axis=1)
            
            st.success("데이터 분석 완료!")
            st.dataframe(df)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            
            st.download_button("📥 결과 파일 다운로드", output.getvalue(), "분석결과.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"계산 중 오류 발생: {e}")
