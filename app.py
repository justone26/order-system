import streamlit as st
import pandas as pd
from io import BytesIO

# 페이지 설정
st.set_page_config(page_title="재고 관리 및 발주 시스템", layout="wide")
st.title("📦 재고 관리 및 발주 시스템")

# 1. 파일 업로드
uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # 파일 형식에 따른 엔진 자동 감지
        if uploaded_file.name.endswith('.xls'):
            df = pd.read_excel(uploaded_file, engine='xlrd')
        else:
            df = pd.read_excel(uploaded_file, engine='openpyxl')
            
        columns = df.columns.tolist()

        def get_default_index(keywords, cols):
            for col in cols:
                for key in keywords:
                    if key in col: return cols.index(col)
            return 0

        # 2. 데이터 매핑 섹션
        st.write("---")
        st.subheader("⚙️ 1단계: 데이터 매핑")
        col1, col2 = st.columns(2)
        with col1:
            sold_out_col = st.selectbox("품절 여부", columns, index=get_default_index(['품절'], columns))
            vendor_col = st.selectbox("공급처", columns, index=get_default_index(['공급처'], columns))
            item_name = st.selectbox("상품명", columns, index=get_default_index(['상품'], columns))
            option_name = st.selectbox("옵션", columns, index=get_default_index(['옵션'], columns))
            vendor_option = st.selectbox("공급처옵션", columns, index=get_default_index(['공급처옵션'], columns))
        with col2:
            stock_col = st.selectbox("정상재고", columns, index=get_default_index(['정상'], columns))
            avail_col = st.selectbox("가용재고", columns, index=get_default_index(['가용'], columns))
            target_3day = st.selectbox("3일 발주 합계", columns, index=get_default_index(['3일'], columns))
            target_1week = st.selectbox("1주 발주 합계", columns, index=get_default_index(['1주', '7일'], columns))

        # 3. 기간 기반 산출 설정
        st.write("---")
        st.subheader("⚙️ 2단계: 기간 기반 산출 설정")
        col3, col4 = st.columns(2)
        with col3:
            lead_time_days = st.number_input("평균 리드타임 기간 (일)", min_value=0, value=0)
        with col4:
            safety_stock_days = st.number_input("안전재고 확보 기간 (판매량의 몇 일분?)", min_value=0, value=3)

        # 4. 분석 실행
        if st.button("🚀 분석 실행"):
            # 수치 데이터 정제
            for col in [target_3day, target_1week, avail_col]:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            # 계산 로직
            df['일일 판매량(기준)'] = (df[target_3day] / 3).round(0).astype(int)
            df['리드타임 준비량'] = (df['일일 판매량(기준)'] * lead_time_days).astype(int)
            df['안전재고 수량'] = (df['일일 판매량(기준)'] * safety_stock_days).astype(int)
            
            # 권장 발주량: (리드타임 준비량 + 안전재고 수량) - 가용재고
            df['권장 발주량'] = (df['리드타임 준비량'] + df['안전재고 수량'] - df[avail_col]).clip(lower=0).astype(int)
            
            df['상태'] = df.apply(lambda row: '🚨 품절/긴급' if (str(row[sold_out_col]).upper() == 'Y' or row[avail_col] <= 0) else '정상', axis=1)

            st.subheader("📊 분석 결과")
            st.dataframe(df)

            # 5. 엑셀 다운로드 기능
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='분석결과')
            
            st.download_button(
                label="📥 분석 결과 엑셀 다운로드",
                data=buffer.getvalue(),
                file_name="재고_분석_결과.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"데이터 처리 중 오류 발생: {e}")
