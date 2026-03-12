import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(layout="wide", page_title="재고 관리 시스템")

# [1] 데이터 저장 및 초기화
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'history' not in st.session_state: st.session_state.history = {}
if 'reorder_db' not in st.session_state: st.session_state.reorder_db = {} 
if 'prev_avail' not in st.session_state: st.session_state.prev_avail = {}

# 구글 시트 연결 함수 (Secrets 설정 필수)
def get_sheet():
    creds_dict = st.secrets["gcp_service_account"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"])
    client = gspread.authorize(creds)
    return client.open("재고관리시트").sheet1

st.title("📦 재고 관리 및 발주 시스템")

# [데이터 로드]
if st.button("🔄 구글 시트에서 최신 데이터 가져오기"):
    try:
        sheet = get_sheet()
        data = sheet.get_all_records()
        df = pd.DataFrame(data)
        st.session_state.df_raw = df
        st.success("시트 로드 완료!")
        st.rerun()
    except Exception as e: st.error(f"시트 연결 오류: {e}")

if st.session_state.df_raw is not None:
    df = st.session_state.df_raw

    # [스마트 리오더 로직: 가용재고 증가 감지 시 자동 차감]
    for idx, row in df.iterrows():
        key = f"{row.get('상품명', '')}_{row.get('옵션', '')}"
        current_avail = row.get('가용재고', 0)
        prev = st.session_state.prev_avail.get(key, current_avail)
        
        if current_avail > prev:
            diff = current_avail - prev
            if key in st.session_state.reorder_db:
                st.session_state.reorder_db[key]['1차'] = max(0, st.session_state.reorder_db[key]['1차'] - diff)
        
        st.session_state.prev_avail[key] = current_avail
        
        # 리오더 DB 값을 df에 적용
        if key in st.session_state.reorder_db:
            df.at[idx, "1차 리오더"] = st.session_state.reorder_db[key]['1차']
            df.at[idx, "2차 리오더"] = st.session_state.reorder_db[key]['2차']

    # [4단계: 데이터 편집]
    st.subheader("📊 4단계: 데이터 편집")
    edited_df = st.data_editor(df, use_container_width=True)
    
    if st.button("💾 리오더 수량 시트 및 DB 저장"):
        # 1. 메모리 저장
        for idx, row in edited_df.iterrows():
            key = f"{row.get('상품명', '')}_{row.get('옵션', '')}"
            st.session_state.reorder_db[key] = {'1차': row["1차 리오더"], '2차': row["2차 리오더"]}
        
        # 2. 구글 시트 업데이트
        sheet = get_sheet()
        sheet.clear()
        sheet.update([edited_df.columns.values.tolist()] + edited_df.values.tolist())
        st.success("데이터가 시트와 DB에 안전하게 저장되었습니다!")

    # [5단계: 발주 리스트 요약]
    st.subheader("📋 5단계: 발주 리스트 요약")
    if '권장 발주량' in edited_df.columns:
        to_order = edited_df[edited_df['권장 발주량'] > 0].copy()
        if not to_order.empty:
            st.dataframe(to_order, use_container_width=True)
            if st.button("💾 이 리스트 기록 저장"):
                today, now = datetime.now().strftime("%Y-%m-%d"), datetime.now().strftime("%H:%M:%S")
                if today not in st.session_state.history: st.session_state.history[today] = {}
                st.session_state.history[today][now] = to_order.copy().reset_index(drop=True)
                st.success("저장 완료!")

    # [6단계: 과거 기록]
    st.subheader("📜 6단계: 과거 데이터 확인")
    for date_key in sorted(st.session_state.history.keys(), reverse=True):
        with st.expander(f"📅 {date_key}"):
            for time_key in sorted(st.session_state.history[date_key].keys(), reverse=True):
                if st.button(f"⏰ {time_key} 보기", key=f"{date_key}_{time_key}"):
                    st.dataframe(st.session_state.history[date_key][time_key])
