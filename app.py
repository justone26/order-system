import streamlit as st
import pandas as pd
import sqlite3
from io import BytesIO
from datetime import datetime

# 1. DB 초기화 (영구 저장용)
def init_db():
    conn = sqlite3.connect('inventory.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS history 
                 (date TEXT, vendor TEXT, item TEXT, qty INTEGER, time TEXT)''')
    conn.commit()
    conn.close()

init_db()

st.set_page_config(layout="wide", page_title="재고 관리 시스템")
st.title("📦 재고 관리 및 발주 시스템")

# 2. 파일 업로드
uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'csv'])

if uploaded_file is not None:
    if 'df' not in st.session_state:
        st.session_state.df = pd.read_excel(uploaded_file) if uploaded_file.name.endswith('.xlsx') else pd.read_csv(uploaded_file)
    
    df = st.session_state.df

    # 3. 분석 실행
    if st.button("🚀 분석 실행"):
        # 가용재고 10 미만 시 20개 발주 권장 (예시 로직)
        df['권장발주량'] = df['가용재고'].apply(lambda x: 20 if x < 10 else 0)
        df['상태'] = df['가용재고'].apply(lambda x: '품절/긴급' if x <= 0 else '정상')
        st.session_state.df = df
        st.rerun()

    # 4. 데이터 편집 및 필터링
    st.subheader("📊 4단계: 검색 및 데이터 편집")
    c1, c2 = st.columns(2)
    with c1:
        search = st.text_input("🔍 상품명 검색")
    with c2:
        status_options = st.multiselect("🚫 상태 필터", options=df['상태'].unique(), default=df['상태'].unique())

    df_disp = df[df['상품명'].str.contains(search, na=False)] if search else df
    df_disp = df_disp[df_disp['상태'].isin(status_options)]
    
    edited_df = st.data_editor(df_disp, use_container_width=True)
    st.session_state.df.update(edited_df)

    # 5. 영구 저장 로직
    if st.button("💾 데이터 영구 저장"):
        conn = sqlite3.connect('inventory.db')
        to_save = st.session_state.df[st.session_state.df['권장발주량'] > 0].copy()
        to_save['date'] = datetime.now().strftime("%Y-%m-%d")
        to_save['time'] = datetime.now().strftime("%H:%M:%S")
        to_save.to_sql('history', conn, if_exists='append', index=False)
        conn.close()
        st.success("데이터베이스에 저장되었습니다!")

    # 6. 엑셀 다운로드
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        st.session_state.df.to_excel(writer, index=False)
    st.download_button("📥 최종 데이터 다운로드", data=buffer.getvalue(), file_name="결과.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # 7. 과거 기록 확인
    st.subheader("📜 6단계: 과거 데이터 확인")
    if st.button("🔍 과거 기록 불러오기"):
        conn = sqlite3.connect('inventory.db')
        history_df = pd.read_sql('SELECT * FROM history', conn)
        conn.close()
        st.dataframe(history_df, use_container_width=True)
