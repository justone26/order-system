import streamlit as st
import pandas as pd
import os

# 페이지 설정
st.set_page_config(page_title="재고 관리 및 발주 시스템", layout="wide")
st.title("📦 재고 관리 및 발주 시스템")

# 1. 파일 업로드
uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요", type=['xlsx', 'xls'])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
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
        lead_time_days = st.number_input("평균 리드타임 기간 (일)", min_value=0, value=7)
    with col4:
        safety_stock_days = st.number_input("안전재고 확보 기간 (판매량의 몇 일분?)", min_value=0, value=3)

    # 4. 분석 실행 및 저장
    if st.button("🚀 분석 실행 및 이력 저장"):
        try:
            for col in [target_3day, target_1week, avail_col]:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            # 계산 로직
            df['일일 판매량(기준)'] = (df[target_3day] / 3).round(0).astype(int)
            df['판매 추이(1주대비)'] = (df['일일 판매량(기준)'] / ((df[target_1week] / 7) + 0.1)).round(2)
            df['리드타임 준비량'] = (df['일일 판매량(기준)'] * lead_time_days).astype(int)
            df['안전재고 수량'] = (df['일일 판매량(기준)'] * safety_stock_days).astype(int)
            
            # [요청 적용] 권장 발주량: (준비량 + 안전재고) - 가용재고
            df['권장 발주량'] = (df['리드타임 준비량'] + df['안전재고 수량'] - df[avail_col]).clip(lower=0).astype(int)
            
            df['상태'] = df.apply(lambda row: '🚨 품절/긴급' if (str(row[sold_out_col]).upper() == 'Y' or row[avail_col] <= 0) else '정상', axis=1)
            df['저장날짜'] = pd.Timestamp.now().strftime('%Y-%m-%d')

            st.subheader("📊 분석 결과")
            st.dataframe(df)

            # 이력 저장
            df.to_csv('order_history.csv', mode='a', header=not os.path.exists('order_history.csv'), index=False)
            st.success("데이터가 이력에 저장되었습니다.")
        except Exception as e:
            st.error(f"분석 오류: {e}")

    # 5. 과거 데이터 조회 섹션
    st.write("---")
    st.subheader("📅 과거 데이터 및 변화량 조회")
    if os.path.exists('order_history.csv'):
        try:
            history_df = pd.read_csv('order_history.csv')
            if '저장날짜' in history_df.columns:
                selected_date = st.selectbox("조회할 날짜를 선택하세요", options=reversed(history_df['저장날짜'].unique()))
                filtered_history = history_df[history_df['저장날짜'] == selected_date]
                st.dataframe(filtered_history)
        except Exception:
            st.info("이력 데이터를 불러오는 중 오류가 발생했습니다.")
    else:
        st.info("저장된 이력이 없습니다. '분석 실행' 버튼을 눌러 첫 데이터를 저장하세요.")
