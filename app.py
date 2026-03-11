import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="재고 관리 및 발주 시스템", layout="wide")
st.title("📦 재고 관리 및 발주 시스템")

uploaded_file = st.file_uploader("엑셀 또는 CSV 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

if uploaded_file is not None:
    try:
        # 파일 로드 및 중복 컬럼 자동 제거
        if uploaded_file.name.endswith('.csv'):
            try: df = pd.read_csv(uploaded_file, encoding='utf-8')
            except: df = pd.read_csv(uploaded_file, encoding='cp949')
        elif uploaded_file.name.endswith('.xls'):
            df = pd.read_excel(uploaded_file, engine='xlrd')
        else:
            df = pd.read_excel(uploaded_file, engine='openpyxl')
        
        df = df.loc[:, ~df.columns.duplicated()]
        columns = df.columns.tolist()

        # 자동 매핑 함수
        def get_best_match(keywords, cols):
            for key in keywords:
                for idx, col in enumerate(cols):
                    if key.lower() in str(col).lower().replace(" ", ""):
                        return idx
            return 0

        # 매핑 설정
        st.write("---")
        st.subheader("⚙️ 1단계: 자동 매핑 확인")
        col1, col2 = st.columns(2)
        with col1:
            sold_out_col = st.selectbox("품절 여부", columns, index=get_best_match(['품절', '판매중단'], columns))
            vendor_col = st.selectbox("공급처", columns, index=get_best_match(['공급처', '업체명'], columns))
            item_name = st.selectbox("상품명", columns, index=get_best_match(['상품명', '상품'], columns))
            option_name = st.selectbox("옵션", columns, index=get_best_match(['옵션'], columns))
        with col2:
            stock_col = st.selectbox("정상재고", columns, index=get_best_match(['정상재고', '재고'], columns))
            avail_col = st.selectbox("가용재고", columns, index=get_best_match(['가용재고', '가용'], columns))
            target_3day = st.selectbox("3일 발주 합계", columns, index=get_best_match(['3일', '최근3일'], columns))
            target_1week = st.selectbox("1주 발주 합계", columns, index=get_best_match(['1주', '7일', '최근7일'], columns))

        # [수정] 엑셀에 리오더 컬럼이 없으면 프로그램 내에서 생성
        pending_col = "입고예정수량(리오더)"
        if pending_col not in df.columns:
            df[pending_col] = 0

        st.write("---")
        st.subheader("⚙️ 2단계: 기간 설정")
        c1, c2 = st.columns(2)
        with c1: lead_time_days = st.number_input("평균 리드타임 (일)", min_value=0, value=0)
        with c2: safety_stock_days = st.number_input("안전재고 확보 기간 (일)", min_value=0, value=3)

        if st.button("🚀 분석 실행"):
            # 수치 정제
            for col in [target_3day, target_1week, avail_col, pending_col, stock_col]:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            # 계산
            df['일일 판매량(기준)'] = (df[target_3day] / 3).round(0).astype(int)
            df['리드타임 준비량'] = (df['일일 판매량(기준)'] * lead_time_days).astype(int)
            df['안전재고 수량'] = (df['일일 판매량(기준)'] * safety_stock_days).astype(int)
            
            # 여기서 화면에서 수정된 pending_col을 반영하여 계산
            df['권장 발주량'] = (df['리드타임 준비량'] + df['안전재고 수량'] - (df[avail_col] + df[pending_col])).clip(lower=0).astype(int)
            df['상태'] = df.apply(lambda row: '🚨 품절/긴급' if (str(row[sold_out_col]).upper() == 'Y' or row[avail_col] <= 0) else '정상', axis=1)

            # 출력용 컬럼 설정
            output_cols = [sold_out_col, vendor_col, item_name, option_name, stock_col, avail_col, pending_col, target_3day, target_1week, '권장 발주량', '상태']
            result_df = df[[c for c in output_cols if c in df.columns]]

            
            st.subheader("📊 분석 결과 (입고예정수량 수정 가능)")
            
            # 여기서 수정하면 edited_df에 반영됨
            edited_df = st.data_editor(result_df, use_container_width=True)

            # 수정된 값을 반영하여 엑셀 저장
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                edited_df.to_excel(writer, index=False, sheet_name='분석결과')
            
            st.download_button("📥 수정된 결과 엑셀 다운로드", data=buffer.getvalue(), file_name="최종_발주서.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"오류 발생: {e}")
