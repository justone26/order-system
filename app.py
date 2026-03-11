import streamlit as st
import pandas as pd
from io import BytesIO

# 페이지 설정
st.set_page_config(page_title="재고 관리 및 발주 시스템", layout="wide")
st.title("📦 재고 관리 및 발주 시스템")

# 1. 파일 업로드
uploaded_file = st.file_uploader("엑셀 또는 CSV 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

if uploaded_file is not None:
    try:
        # 파일 형식 자동 감지 및 로드
        file_name = uploaded_file.name.lower()
        if file_name.endswith('.csv'):
            try: df = pd.read_csv(uploaded_file, encoding='utf-8')
            except: df = pd.read_csv(uploaded_file, encoding='cp949')
        elif file_name.endswith('.xls'):
            df = pd.read_excel(uploaded_file, engine='xlrd')
        else:
            df = pd.read_excel(uploaded_file, engine='openpyxl')
            
        columns = df.columns.tolist()

        # 자동 매핑 함수
        def get_best_match(keywords, cols):
            for key in keywords:
                for idx, col in enumerate(cols):
                    if key.lower() in str(col).lower().replace(" ", ""):
                        return idx
            return 0

        # 2. 데이터 매핑 섹션
        st.write("---")
        st.subheader("⚙️ 1단계: 자동 매핑 확인")
        col1, col2 = st.columns(2)
        with col1:
            sold_out_col = st.selectbox("품절 여부", columns, index=get_best_match(['품절', '판매중단'], columns))
            vendor_col = st.selectbox("공급처", columns, index=get_best_match(['공급처', '업체명'], columns))
            item_name = st.selectbox("상품명", columns, index=get_best_match(['상품명', '상품'], columns))
            option_name = st.selectbox("옵션", columns, index=get_best_match(['옵션'], columns))
            vendor_option = st.selectbox("공급처옵션", columns, index=get_best_match(['공급처옵션', '거래처옵션'], columns))
        with col2:
            stock_col = st.selectbox("정상재고", columns, index=get_best_match(['정상재고', '재고'], columns))
            avail_col = st.selectbox("가용재고", columns, index=get_best_match(['가용재고', '가용'], columns))
            pending_col = st.selectbox("입고예정수량(리오더)", columns, index=get_best_match(['입고예정', '리오더', '입고대기'], columns))
            target_3day = st.selectbox("3일 발주 합계", columns, index=get_best_match(['3일', '최근3일'], columns))
            target_1week = st.selectbox("1주 발주 합계", columns, index=get_best_match(['1주', '7일', '최근7일'], columns))

        # 3. 기간 설정
        st.write("---")
        st.subheader("⚙️ 2단계: 기간 설정")
        c1, c2 = st.columns(2)
        with c1: lead_time_days = st.number_input("평균 리드타임 (일)", min_value=0, value=0)
        with c2: safety_stock_days = st.number_input("안전재고 확보 기간 (일)", min_value=0, value=3)

        # 4. 분석 실행
        if st.button("🚀 분석 실행"):
            # 수치 데이터 정제
            for col in [target_3day, target_1week, avail_col, pending_col, stock_col]:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            # 계산 로직 (리오더 수량 고려)
            df['일일 판매량(기준)'] = (df[target_3day] / 3).round(0).astype(int)
            df['리드타임 준비량'] = (df['일일 판매량(기준)'] * lead_time_days).astype(int)
            df['안전재고 수량'] = (df['일일 판매량(기준)'] * safety_stock_days).astype(int)
            
            # [핵심 로직] 권장 발주량 = (필요량) - (현재 가용재고 + 이미 발주된 입고예정수량)
            df['권장 발주량'] = (df['리드타임 준비량'] + df['안전재고 수량'] - (df[avail_col] + df[pending_col])).clip(lower=0).astype(int)
            df['상태'] = df.apply(lambda row: '🚨 품절/긴급' if (str(row[sold_out_col]).upper() == 'Y' or row[avail_col] <= 0) else '정상', axis=1)

            # 필수 컬럼 리스트 구성
            output_cols = [
                sold_out_col, vendor_col, item_name, option_name, vendor_option, 
                stock_col, avail_col, pending_col, target_3day, target_1week, 
                '일일 판매량(기준)', '리드타임 준비량', '안전재고 수량', '권장 발주량', '상태'
            ]
            
            unique_cols = list(dict.fromkeys([c for c in output_cols if c in df.columns]))
            result_df = df[unique_cols]

            
            st.subheader("📊 분석 결과")
            st.dataframe(result_df)

            # 5. 엑셀 다운로드
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False, sheet_name='분석결과')
            
            st.download_button("📥 분석 결과 엑셀 다운로드", data=buffer.getvalue(), file_name="재고_분석_결과.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"오류 발생: {e}")
