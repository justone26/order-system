import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# [강력 조치] 소스 수정 시 화면 갱신을 돕기 위한 버전 관리 (숫자를 바꾸면 강제 리셋됨)
VERSION = "1.2" 

# 1. 기본 설정
KST = timezone(timedelta(hours=9))
st.set_page_config(layout="wide", page_title=f"저스트원 재고관리 v{VERSION}")

# 2단계 자동 매칭 10가지 키워드 (사장님 요청사항)
MATCH_KEYS = {
    "상품명": ["상품명", "물명", "아이템", "Item Name", "상품", "제품명"],
    "옵션": ["옵션", "규격", "사이즈", "컬러", "Option", "색상"],
    "가용재고": ["가용재고", "현재고", "재고수량", "재고", "Stock", "실재고"],
    "7일판매량": ["7일판매량", "주간판매", "7일판매", "판매량(7일)", "Sale7", "최근판매"],
    "원가": ["원가", "매입가", "구매가", "Cost", "단가"],
    "판매가": ["판매가", "매출가", "정가", "Price", "소비자가"],
    "바코드": ["바코드", "품번", "관리코드", "Barcode", "SKU"],
    "거래처": ["거래처", "제조사", "공급처", "Vendor", "처명"],
    "상태": ["상태", "판매상태", "진열상태", "Status", "구분"],
    "카테고리": ["카테고리", "분류", "Category", "대분류"]
}

def auto_match(cols, target_key):
    for i, col in enumerate(cols):
        if any(k in str(col) for k in MATCH_KEYS[target_key]):
            return i
    return 0

# 구글 시트 함수들 생략 없이 유지...
def get_sheet():
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        return client.open_by_key("1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg")
    except: return None

# 세션 관리
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'df_final' not in st.session_state: st.session_state.df_final = None

st.title(f"📦 저스트원 통합 재고 관리 v{VERSION}")

# 1단계: 업로드
up_file = st.file_uploader("📁 1단계: 엑셀 업로드", type=['xlsx', 'xls', 'csv'])

if up_file:
    # 파일이 새로 올라오면 무조건 초기화
    df = pd.read_csv(up_file) if up_file.name.endswith('.csv') else pd.read_excel(up_file)
    df.columns = df.columns.str.strip()
    st.session_state.df_raw = df
    
    # --- 2단계: 자동 매칭 노출 ---
    st.divider()
    st.subheader("🔗 2단계: 자동 컬럼 매칭 (10종)")
    cols = df.columns.tolist()
    
    c1, c2, c3, c4 = st.columns(4)
    with c1: s_item = st.selectbox("상품명", cols, index=auto_match(cols, "상품명"))
    with c2: s_opt = st.selectbox("옵션", cols, index=auto_match(cols, "옵션"))
    with c3: s_avail = st.selectbox("가용재고", cols, index=auto_match(cols, "가용재고"))
    with c4: s_t7 = st.selectbox("7일판매량", cols, index=auto_match(cols, "7일판매량"))
    
    with st.expander("기타 매칭 항목 (원가, 거래처 등)"):
        c5, c6, c7 = st.columns(3)
        with c5: st.selectbox("원가", cols, index=auto_match(cols, "원가"))
        with c6: st.selectbox("거래처", cols, index=auto_match(cols, "거래처"))
        with c7: st.selectbox("바코드", cols, index=auto_match(cols, "바코드"))

    # --- 3단계: 분석 버튼 ---
    st.divider()
    if st.button("🚀 3단계: 분석 실행", use_container_width=True):
        # 분석 로직... (이하 동일)
        st.session_state.df_final = df # 결과 저장
        st.rerun()

# 4~6단계는 df_final이 있을 때만 아래에 표시... (생략)
