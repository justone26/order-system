import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import holidays

# 페이지 설정
st.set_page_config(layout="wide", page_title="재고 관리 시스템")

# 1. 디자인: 제목 및 스타일
st.markdown("""
    <div style="font-size: 55px; font-weight: 900; color: #000; margin-bottom: 20px;">
        📦 재고 관리 및 발주 시스템
    </div>
""", unsafe_allow_html=True)

if st.button("🔄 시스템 초기화"):
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

# 세션 관리 및 로직 함수
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'history' not in st.session_state: st.session_state.history = {}

def merge_new_data(old_df, new_df):
    if old_df is None: return new_df
    old_df['key'] = old_df['상품명'].astype(str) + "_" + old_df['옵션'].astype(str)
    new_df['key'] = new_df['상품명'].astype(str) + "_" + new_df['옵션'].astype(str)
    if '1차 입고예정' not in old_df.columns: old_df['1차 입고예정'] = 0
    if '2차 입고예정' not in old_df.columns: old_df['2차 입고예정'] = 0
    manual_data = old_df[['key', '1차 입고예정', '2차 입고예정']]
    merged_df = pd.merge(new_df, manual_data, on='key', how='left')
    merged_df['1차 입고예정'] = merged_df['1차 입고예정'].fillna(0)
    merged_df['2차 입고예정'] = merged_df['2차 입고예정'].fillna(0)
    return merged_df.drop(columns=['key'])

def auto_reduce_reorder(old_df, edited_df):
    diff = edited_df['가용재고'] - old_df['가용재고']
    for i in diff[diff > 0].index:
        amount = diff[i]
        reduce_1 = min(edited_df.at[i, '1차 입고예정'], amount)
        edited_df.at[i, '1차 입고예정'] -= reduce_1
        amount -= reduce_1
        if amount > 0:
            reduce_2 = min(edited_df.at[i, '2차 입고예정'], amount)
            edited_df.at[i, '2차 입고예정'] -= reduce_2
    return edited_df

# 1~6단계 메인 로직
uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'xls', 'csv'])
if uploaded_file is not None:
    new_df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
    st.session_state.df_raw = merge_new_data(st.session_state.df_raw, new_df)
    st.rerun()

if st.session_state.df_raw is not None:
    # 매핑 및 분석 단계 (생략된 부분은 이전 코드와 동일하게 붙여넣으세요)
    # ... (데이터 편집 단계) ...
    st.subheader("📊 4단계: 검색 및 데이터 편집")
    df_final = st.session_state.df_raw[['상품명', '옵션', '가용재고', '1차 입고예정', '2차 입고예정', '권장 발주량']]
    edited_df = st.data_editor(df_final, use_container_width=True)
    
    if not edited_df.equals(df_final):
        updated_df = auto_reduce_reorder(df_final, edited_df)
        st.session_state.df_raw.update(updated_df)
        st.rerun()

    # 5단계 요약 및 6단계 기록 등 기존 로직 모두 포함 가능
    st.subheader("📋 5단계: 발주 요약 및 기록")
    if st.button("💾 데이터 및 히스토리 저장"):
        date_key = datetime.now().strftime("%Y-%m-%d")
        if date_key not in st.session_state.history: st.session_state.history[date_key] = []
        st.session_state.history[date_key].append(st.session_state.df_raw.copy())
        st.success("저장 완료!")

    st.subheader("📜 6단계: 과거 기록")
    if st.session_state.history:
        date_sel = st.selectbox("날짜 선택", list(st.session_state.history.keys()))
        st.dataframe(st.session_state.history[date_sel][-1])
