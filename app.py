import streamlit as st
import pandas as pd
import os
import io

st.set_page_config(page_title="자동 발주 및 재고 관리", layout="wide")
st.title("📦 재고 품절 방지 및 발주 관리 시스템")

# 세션 상태 초기화
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
    st.subheader("⚙️ 데이터 항목 매핑")
    col1, col2 = st.columns(2)
    with col1:
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

    if st.button("분석 및 이력 저장"):
        try:
            df['일 판매 데이터'] = df[col_invoice] + df[col_reception]
            df['일일평균'] = df[target_3day] / 3
            df['재고소진일'] = (df[stock_col] / df['일일평균'].replace(0, 1)).round(1)
            df['상태'] = df.apply(lambda row: '🚨 품절/긴급' if (str(row[sold_out_col]).upper() == 'Y' or row[stock_col] <= 0 or row['재고소진일'] < 3) else '정상', axis=1)
            df['저장날짜'] = pd.Timestamp.now().strftime('%Y-%m-%d')
            df.to_csv('order_history.csv', mode='a', header=not os.path.exists('order_history.csv'), index=False)
            st.session_state.analysis_result = df
            st.success("데이터 분석 및 이력 저장 완료!")
        except Exception as e:
            st.error(f"오류 발생: {e}")

    if st.session_state.analysis_result is not None:
        st.subheader("📊 최신 분석 결과")
        st.dataframe(st.session_state.analysis_result)
        
        # 다운로드 버튼
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            st.session_state.analysis_result.to_excel(writer, index=False)
        st.download_button("📥 결과 파일 다운로드 (Excel)", buffer.getvalue(), "최신_분석결과.xlsx")

    st.write("---")
    st.subheader("📅 과거 발주 내역 검색 및 추이 비교")
    search_date = st.date_input("조회할 날짜 선택")
    
    if st.button("내역 조회"):
        if os.path.exists('order_history.csv'):
            history_df = pd.read_csv('order_history.csv')
            filtered_df = history_df[history_df['저장날짜'] == str(search_date)]
            
            if st.session_state.analysis_result is not None and not filtered_df.empty:
                st.write("📈 **옵션별 판매량 추이 비교**")
                compare_df = pd.merge(
                    st.session_state.analysis_result[[item_name, option_name, '일 판매 데이터']], 
                    filtered_df[[item_name, option_name, '일 판매 데이터']], 
                    on=[item_name, option_name], suffixes=('_현재', '_과거')
                )
                compare_df['판매량 변화'] = compare_df['일 판매 데이터_현재'] - compare_df['일 판매 데이터_과거']
                
                # 시각화
                st.bar_chart(compare_df.set_index(option_name)['판매량 변화'])
                
                # 색상 적용 스타일링
                def color_negative_red(val):
                    color = 'red' if val > 0 else 'blue' if val < 0 else 'black'
                    return f'color: {color}'
                
                styled_df = compare_df.sort_values(by='판매량 변화', ascending=False).style.applymap(color_negative_red, subset=['판매량 변화'])
                st.dataframe(styled_df)
            elif filtered_df.empty:
                st.warning("해당 날짜에 저장된 내역이 없습니다.")
        else:
            st.error("저장된 이력이 없습니다.")
