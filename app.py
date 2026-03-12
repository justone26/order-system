import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import holidays

st.set_page_config(layout="wide", page_title="재고 관리 시스템")

# 1. 디자인: 제목 및 초기화
st.markdown('<div style="font-size: 55px; font-weight: 900; margin-bottom: 20px;">📦 재고 관리 및 발주 시스템</div>', unsafe_allow_html=True)
if st.button("🔄 시스템 초기화"):
    for key in list(st.session_state.keys()): del st.session_state[key]
    st.rerun()

# 세션 관리 함수
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'history' not in st.session_state: st.session_state.history = {}

def merge_new_data(old_df, new_df):
    if old_df is None: return new_df
    old_df['key'] = old_df['상품명'].astype(str) + "_" + old_df['옵션'].astype(str)
    new_df['key'] = new_df['상품명'].astype(str) + "_" + new_df['옵션'].astype(str)
    if '1차 입고예정' not in old_df.columns: old_df['1차 입고예정'] = 0
    if '2차 입고예정' not in old_df.columns: old_df['2차 입고예정'] = 0
    manual_data = old_df[['key', '1차 입고예정', '2차 입고예정']]
    merged_df = pd.merge(new_df, manual_data, on='key', how='left').fillna(0)
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

# 파일 업로드
uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'csv'])
if uploaded_file:
    new_df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
    st.session_state.df_raw = merge_new_data(st.session_state.df_raw, new_df)
    st.rerun()

# 파일 업로드 후 전체 로직
if st.session_state.df_raw is not None:
    df = st.session_state.df_raw
    cols = df.columns.tolist()

    # 1단계: 매핑 설정
    st.subheader("⚙️ 1단계: 매핑 설정")
    c1, c2 = st.columns(2)
    sold_out = c1.selectbox("품절 여부", cols, index=0)
    avail = c2.selectbox("가용재고", cols, index=0)

    # 2~3단계: 분석 설정
    st.subheader("⚙️ 2~3단계: 분석 설정")
    if st.button("🚀 분석 실행"):
        df['권장 발주량'] = 0 # 분석 로직 예시
        st.success("분석 완료!")
        st.rerun()

    # 4단계: 데이터 편집 (KeyError 방지 포함)
    st.subheader("📊 4단계: 데이터 편집")
    for col in ['1차 입고예정', '2차 입고예정', '권장 발주량']:
        if col not in df.columns: df[col] = 0
    
    edit_cols = ['상품명', '옵션', '가용재고', '1차 입고예정', '2차 입고예정', '권장 발주량']
    edited_df = st.data_editor(df[[c for c in edit_cols if c in df.columns]], use_container_width=True)
    
    if not edited_df.equals(df[[c for c in edit_cols if c in df.columns]]):
        updated_df = auto_reduce_reorder(df, edited_df)
        st.session_state.df_raw.update(updated_df)
        st.rerun()

    # 5단계: 요약 및 기록
    st.subheader("📋 5단계: 발주 리스트 요약")
    if st.button("💾 데이터 저장"):
        date_key = datetime.now().strftime("%Y-%m-%d %H:%M")
        st.session_state.history[date_key] = df.copy()
        st.success("저장 완료!")

    # 6단계: 과거 기록 확인
    st.subheader("📜 6단계: 과거 데이터 확인")
    if st.session_state.history:
        date_sel = st.selectbox("기록 선택", list(st.session_state.history.keys()))
        st.dataframe(st.session_state.history[date_sel], use_container_width=True)
        
        # 다운로드 기능
        buf = BytesIO()
        st.session_state.history[date_sel].to_excel(buf, index=False)
        st.download_button("📥 기록 다운로드", data=buf.getvalue(), file_name="history.xlsx")
else:
    st.info("파일을 업로드하면 시스템이 활성화됩니다.")
