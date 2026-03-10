import streamlit as st
import pandas as pd
import os
import io

st.set_page_config(page_title="자동 발주 및 재고 관리", layout="wide")
st.title("📦 리드타임/안전재고 반영 발주 관리 시스템")

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
    col1, col2 = st.columns(2)
    with col1:
        vendor_name = st.selectbox("공급처", columns, index=get_default_index(['공급처'], columns))
        item_name = st.selectbox("상품명", columns, index=get_default_index(['상품'], columns))
        option_name = st.selectbox("옵션", columns, index=get_default_index(['옵션'], columns))
        stock_col = st.selectbox("정상재고", columns, index=get_default_index(['재고'], columns))
        available_stock = st.selectbox("가용재고", columns, index=get_default_index(['가용'], columns))
    with col2:
        col_invoice = st.selectbox("송장", columns, index=get_default_index(['송장'], columns))
        col_reception = st.selectbox("접수", columns, index=get_default_index(['접수'], columns))
        target_3day = st.selectbox("3일 발주 합계", columns, index=get_default_index(['3일'], columns))
        lead_time = st.number_input("평균 리드타임 (일)", min_value=0, value=7)
        safety_stock_days = st.number_input("안전재고 확보일 (일)", min_value=0, value=3)

    # 1. 분석 및 이력 저장
    if st.button("분석 및 이력 저장"):
        try:
            for col in [col_invoice, col_reception, target_3day, stock_col, available_stock]:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
            df['일 판매 데이터'] = df[col_invoice] + df[col_reception]
            df['일일평균'] = (df[target_3day] / 3).round(1)
            df['권장발주수량'] = (df['일일평균'] * (lead_time + safety_stock_days) - df[available_stock]).apply(lambda x: max(0, int(x)))
            df['저장날짜'] = pd.Timestamp.now().strftime('%Y-%m-%d')
            
            # CSV 저장
            df.to_csv('order_history.csv', mode='a', header=not os.path.exists('order_history.csv'), index=False)
            st.session_state.analysis_result = df
            st.success("분석 완료 및 이력 저장 성공!")
        except Exception as e:
            st.error(f"오류: {e}")

    # 2. 결과 출력
    if st.session_state.analysis_result is not None:
        st.subheader("📊 최신 분석 결과")
        st.dataframe(st.session_state.analysis_result)

    # 3. 과거 내역 조회 및 비교
    st.write("---")
    st.subheader("📅 과거 발주 내역 검색 및 추이 비교")
    search_date = st.date_input("조회할 날짜 선택")
    
    if st.button("내역 조회"):
        if os.path.exists('order_history.csv'):
            history_df = pd.read_csv('order_history.csv')
            filtered_df = history_df[history_df['저장날짜'] == str(search_date)]
            
            if not filtered_df.empty and st.session_state.analysis_result is not None:
                st.write(f"📈 **{search_date} 데이터와 비교**")
                compare_df = pd.merge(
                    st.session_state.analysis_result[[item_name, option_name, '일 판매 데이터']], 
                    filtered_df[[item_name, option_name, '일 판매 데이터']], 
                    on=[item_name, option_name], suffixes=('_현재', '_과거')
                )
                compare_df['판매량 변화'] = compare_df['일 판매 데이터_현재'] - compare_df['일 판매 데이터_과거']
                
                
                st.bar_chart(compare_df.set_index(option_name)['판매량 변화'])
                
                def color_red_blue(val):
                    return f'color: {"red" if val > 0 else "blue" if val < 0 else "black"}'
                
                styled_df = compare_df.style.applymap(color_red_blue, subset=['판매량 변화'])
                st.dataframe(styled_df)
            else:
                st.warning("선택한 날짜에 저장된 기록이 없습니다.")
        else:
            st.error("저장된 이력 파일이 없습니다.")
