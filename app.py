import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import io

# [기존 함수들 생략 - 동일함]

st.set_page_config(layout="wide", page_title="제작상품 재고관리")

# --- 버튼 왼쪽 정렬을 위한 스타일 추가 ---
st.markdown("""
    <style>
    /* 모든 버튼의 텍스트를 왼쪽으로 정렬 */
    div.stButton > button {
        display: flex;
        justify-content: flex-start;
        padding-left: 15px;
    }
    </style>
    """, unsafe_allow_html=True)

st.title("🏭 제작 상품 재고 관리 시스템")

if "extra_order_dict" not in st.session_state:
    st.session_state.extra_order_dict = {}
if 'analyzed' not in st.session_state:
    st.session_state.analyzed = False

tab1, tab2 = st.tabs(["📊 제작 상품 관리", "📜 히스토리 조회"])

with tab1:
    c_up1, c_up2 = st.columns([3, 1])
    uploaded_file = c_up1.file_uploader("엑셀 파일을 올려주세요", type=['xlsx', 'xls', 'csv'])
    
    # 아이콘을 제거하고 싶으시면 아래 "🔄 " 부분을 지우시면 됩니다.
    if c_up2.button("6단계: 이전 데이터 로드", use_container_width=True):
        try:
            sheet = get_sheet().sheet1
            gs_df = pd.DataFrame(sheet.get_all_records())
            if not gs_df.empty:
                st.session_state.df_raw = gs_df.rename(columns={'상품명': '기존상품명', '옵션': '기존옵션'})
                st.success("데이터 로드 완료")
            else:
                st.warning("데이터 없음")
        except:
            st.error("연결 오류")

# [나머지 코드 생략 - 동일함]
