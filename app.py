import streamlit as st
import pandas as pd

st.set_page_config(page_title="재고 관리 및 판매 추이 분석", layout="wide")
st.title("📦 지능형 재고 관리 시스템 (판매 추이 포함)")

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

    st.write("---")
    st.subheader("⚙️ 2단계: 기간 기반 산출 설정")
    col3, col4 = st.columns(2)
    with col3:
        lead_time_days = st.number_input("평균 리드타임 기간 (일)", min_value=0, value=7)
    with col4:
        safety_stock_days = st.number_input("안전재고 확보 기간 (일)", min_value=0, value=3)

    if st.button("분석 실행"):
        try:
            # 수치 데이터 정제
            for col in [stock_col, avail_col, target_3day, target_1week]:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            # 1. 판매량 산출 및 과거 데이터 대조
            df['일일 판매량(최근 3일)'] = (df[target_3day] / 3).round(1)
            df['일일 판매량(지난 1주)'] = (df[target_1week] / 7).round(1)
            
            # 2. 변화량(추이) 계산: 최근 3일 판매가 1주 평균보다 얼마나 늘었는가
            # 1.0보다 크면 판매 증가, 작으면 판매 감소
            df['판매 변화율(추이)'] = (df['일일 판매량(최근 3일)'] / (df['일일 판매량(지난 1주)'] + 0.1)).round(2)
            
            # 3. 리드타임 및 안전재고 수량 산출
            df['리드타임 준비량'] = (df['일일 판매량(최근 3일)'] * lead_time_days).astype(int)
            df['안전재고 수량'] = (df['일일 판매량(최근 3일)'] * safety_stock_days).astype(int)
            
            # 4. 최종 권장 발주량: (준비량 + 안전재고) - 가용재고
            df['권장 발주량'] = (df['리드타임 준비량'] + df['안전재고 수량'] - df[avail_col]).clip(lower=0).astype(int)

            st.subheader("📊 1. 전체 분석 결과 (과거 데이터 & 변화량 포함)")
            # 보기 편하게 컬럼 순서 재배치
            final_cols = [
                vendor_col, item_name, option_name, 
                target_1week, target_3day, '판매 변화율(추이)',
                '일일 판매량(최근 3일)', '리드타임 준비량', '안전재고 수량', 
                avail_col, '권장 발주량'
            ]
            st.dataframe(df[final_cols])

            st.write("---")
            st.subheader("📈 2. 판매 추이 시각화 (최근 3일 vs 지난 1주)")
            # 상위 15개 상품의 변화량을 그래프로 표시
            chart_data = df.nlargest(15, target_3day)[[option_name, '일일 판매량(지난 1주)', '일일 판매량(최근 3일)']]
            chart_data = chart_data.set_index(option_name)
            st.bar_chart(chart_data)
            
            

        except Exception as e:
            st.error(f"분석 중 오류 발생: {e}")
