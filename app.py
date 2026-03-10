import streamlit as st
import pandas as pd
import plotly.express as px  # 시각화를 위해 추가

st.set_page_config(page_title="재고 관리 및 판매 추이 분석", layout="wide")
st.title("📦 지능형 재고 관리 & 옵션별 판매 분석")

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
        # 기본 계산 로직 (기존 유지)
        df['일일 판매량(기준)'] = (df[target_3day] / 3).round(1)
        df['리드타임 준비량'] = (df['일일 판매량(기준)'] * lead_time_days).astype(int)
        df['안전재고 수량'] = (df['일일 판매량(기준)'] * safety_stock_days).astype(int)
        df['권장 발주량'] = (df['리드타임 준비량'] + df['안전재고 수량'] - df[avail_col]).clip(lower=0).astype(int)
        
        # [추가] 과거 데이터 대비 분석 (판매 추이)
        df['판매 성장률(3일 vs 1주)'] = ((df[target_3day] / 3) / (df[target_1week] / 7 + 0.1)).round(2)

        st.subheader("📊 1. 발주 권장 리스트")
        st.dataframe(df)

        st.write("---")
        st.subheader("📈 2. 옵션별 판매 추이 및 비중 분석")
        
        chart_col1, chart_col2 = st.columns(2)
        
        with chart_col1:
            st.write("**상품별/옵션별 3일 판매 비중**")
            # 
            fig_pie = px.sunburst(df, path=[vendor_col, item_name, option_name], values=target_3day,
                                 color=target_3day, color_continuous_scale='RdBu')
            st.plotly_chart(fig_pie, use_container_width=True)

        with chart_col2:
            st.write("**옵션별 재고 부족 위험도 (가용재고 vs 권장발주)**")
            # 
            fig_bar = px.bar(df.head(20), x=option_name, y=[avail_col, '권장 발주량'], 
                            barmode='group', title="상위 20개 옵션 분석")
            st.plotly_chart(fig_bar, use_container_width=True)

        st.write("---")
        st.subheader("📋 3. 과거 데이터 상세 대조 (3일 vs 1주)")
        # 1주 데이터와 3일 데이터를 비교하여 판매가 늘고 있는지 줄고 있는지 보여줌
        comparison_df = df[[item_name, option_name, target_1week, target_3day, '판매 성장률(3일 vs 1주)']]
        st.table(comparison_df.sort_values(by='판매 성장률(3일 vs 1주)', ascending=False).head(10))
