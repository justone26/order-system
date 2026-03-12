import streamlit as st
import pandas as pd
import pickle
import os
from io import BytesIO
from datetime import datetime

st.set_page_config(layout="wide", page_title="재고 관리 시스템")

# [1] 데이터 저장 및 초기화
SAVE_FILE = "reorder_data.pkl"
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'history' not in st.session_state: st.session_state.history = {}

# 리오더 DB 불러오기 (엑셀을 새로 올려도 값이 유지되게 함)
if 'reorder_db' not in st.session_state:
    if os.path.exists(SAVE_FILE):
        with open(SAVE_FILE, 'rb') as f:
            st.session_state.reorder_db = pickle.load(f)
    else:
        st.session_state.reorder_db = {}

st.title("📦 재고 관리 및 발주 시스템")

# [2] 엑셀 업로드
uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'xls', 'csv'])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
    df = df.loc[:, ~df.columns.duplicated()]
    
    # 리오더 값이 없으면 생성
    if "1차 리오더" not in df.columns: df["1차 리오더"] = 0
    if "2차 리오더" not in df.columns: df["2차 리오더"] = 0
    
    # 리오더 DB 매핑 (기존 저장값 불러오기)
    for idx, row in df.iterrows():
        key = f"{row.get('상품명', '')}_{row.get('옵션', '')}"
        if key in st.session_state.reorder_db:
            df.at[idx, "1차 리오더"] = st.session_state.reorder_db[key]['1차']
            df.at[idx, "2차 리오더"] = st.session_state.reorder_db[key]['2차']
            
    st.session_state.df_raw = df
    st.success("엑셀 파일 로드 완료 및 리오더 데이터 매핑 성공!")

if st.session_state.df_raw is not None:
    # 4단계: 데이터 편집
    st.subheader("📊 4단계: 데이터 편집 및 리오더 저장")
    f1, f2 = st.columns([3, 1])
    filter_mode = f2.selectbox("품절 필터", ["정상만", "품절만", "전체보기"])
    
    df_filtered = st.session_state.df_raw.copy()
    if filter_mode == "정상만": df_filtered = df_filtered[~df_filtered['품절 여부'].astype(str).str.contains('품절', na=False)]
    elif filter_mode == "품절만": df_filtered = df_filtered[df_filtered['품절 여부'].astype(str).str.contains('품절', na=False)]
    
    edited_df = st.data_editor(df_filtered, use_container_width=True)
    
    if st.button("💾 리오더 수치 안전하게 저장"):
        # 편집된 리오더 값을 DB(pickle)에 저장
        for idx, row in edited_df.iterrows():
            key = f"{row.get('상품명', '')}_{row.get('옵션', '')}"
            st.session_state.reorder_db[key] = {'1차': row["1차 리오더"], '2차': row["2차 리오더"]}
        
        # 파일 저장
        with open(SAVE_FILE, 'wb') as f: pickle.dump(st.session_state.reorder_db, f)
        
        # 원본 데이터 업데이트
        st.session_state.df_raw.update(edited_df)
        st.success("데이터가 저장되었습니다!")

    # 5단계: 발주 리스트 요약
    st.subheader("📋 5단계: 발주 리스트 요약")
    if '권장 발주량' in st.session_state.df_raw.columns:
        to_order = st.session_state.df_raw[st.session_state.df_raw['권장 발주량'] > 0]
        st.dataframe(to_order, use_container_width=True)
        
        if st.button("💾 리스트 기록 저장"):
            today, now = datetime.now().strftime("%Y-%m-%d"), datetime.now().strftime("%H:%M:%S")
            if today not in st.session_state.history: st.session_state.history[today] = {}
            st.session_state.history[today][now] = to_order.copy()
            st.success("기록 저장 완료!")

    # 6단계: 과거 데이터 확인
    st.subheader("📜 6단계: 과거 데이터 확인")
    for date_key in sorted(st.session_state.history.keys(), reverse=True):
        with st.expander(f"📅 {date_key}"):
            for time_key in sorted(st.session_state.history[date_key].keys(), reverse=True):
                if st.button(f"⏰ {time_key} 보기", key=f"{date_key}_{time_key}"):
                    st.dataframe(st.session_state.history[date_key][time_key])
