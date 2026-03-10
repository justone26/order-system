import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="자동 발주 및 재고 관리", layout="wide")
st.title("📦 재고 품절 방지 및 발주 관리 시스템")

# [세션 상태 관리]
if 'analysis_result' not in st.session_state: st.session_state.analysis_result = None
if 'history_result' not in st.session_state: st.session_state.history_result = None

uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요", type=['xlsx', 'xls'])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    columns = df.columns.tolist()

    # (매핑 로직은 동일)
    def get_default_index(keywords, cols):
        for col in cols:
            for key in keywords:
                if key in col: return cols.index(col)
        return 0

    st.subheader("⚙️ 데이터 항목 매핑")
    # ... (기존 매핑 코드 생략) ...

    if st.button("분석 및 이력 저장"):
        # ... (기존 분석 및 저장 로직 동일) ...
        st.success("데이터 분석 및 이력 저장 완료!")

    # [스타일링 함수: 판매량 변화 색상 적용]
    def highlight_changes(val):
        color = 'red' if val > 0 else 'blue' if val < 0 else 'black'
        return f'color: {color}'

    # [과거 내역 검색 및 비교]
    if st.button("내역 조회"):
        if os.path.exists('order_history.csv'):
            history_df = pd.read_csv('order_history.csv')
            filtered_df = history_df[history_df['저장날짜'] == str(st.date_input("조회할 날짜"))]
            
            if st.session_state.analysis_result is not None:
                compare_df = pd.merge(
                    st.session_state.analysis_result[['상품명', '옵션', '일 판매 데이터']], 
                    filtered_df[['상품명', '옵션', '일 판매 데이터']], 
                    on=['상품명', '옵션'], suffixes=('_현재', '_과거')
                )
                compare_df['판매량 변화'] = compare_df['일 판매 데이터_현재'] - compare_df['일 판매 데이터_과거']
                
                # [그래프 및 색상 입힌 표 출력]
                st.write("📈 **옵션별 판매량 추이 비교**")
                st.bar_chart(compare_df.set_index('옵션')['판매량 변화'])
                
                # 데이터프레임 스타일링 (판매량 변화 컬럼에 빨강/파랑 색상 적용)
                styled_df = compare_df.sort_values(by='판매량 변화', ascending=False).style.applymap(
                    highlight_changes, subset=['판매량 변화']
                )
                st.dataframe(styled_df)
