import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 1. 기본 설정
KST = timezone(timedelta(hours=9))
st.set_page_config(layout="wide", page_title="저스트원 재고관리")

# 2. 10가지 자동 매칭 키워드 (사장님 요청)
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

# 3. 세션 상태 관리
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'df_final' not in st.session_state: st.session_state.df_final = None

# 초기화 함수
def reset_all():
    st.session_state.df_raw = None
    st.session_state.df_final = None
    st.rerun()

st.title("📦 저스트원 통합 재고 관리 시스템")

tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 관리"])

with tab1:
    # --- [1단계: 데이터 업로드 영역] ---
    # 사진처럼 한 줄에 제목과 초기화 버튼 배치
    col_title, col_btn = st.columns([3, 1])
    with col_title:
        st.subheader("📁 1단계: 데이터 업로드")
    with col_btn:
        # 사진 속의 '전체 데이터 초기화' 버튼 위치
        if st.button("🔄 전체 데이터 초기화", use_container_width=True):
            reset_all()

    up_file = st.file_uploader("엑셀/CSV 파일을 선택하세요", type=['xlsx', 'xls', 'csv'], key="up_key")
    
    if up_file and st.session_state.df_raw is None:
        try:
            df = pd.read_csv(up_file) if up_file.name.endswith('.csv') else pd.read_excel(up_file)
            df.columns = df.columns.str.strip()
            st.session_state.df_raw = df
            st.rerun()
        except Exception as e:
            st.error(f"파일 오류: {e}")

    # --- [2~3단계: 자동 매칭 및 분석] ---
    if st.session_state.df_raw is not None and st.session_state.df_final is None:
        st.divider()
        st.subheader("🔗 2단계: 자동 컬럼 매칭 (10종)")
        cols = st.session_state.df_raw.columns.tolist()
        
        c1, c2, c3, c4 = st.columns(4)
        with c1: sel_item = st.selectbox("상품명", cols, index=auto_match(cols, "상품명"))
        with c2: sel_opt = st.selectbox("옵션", cols, index=auto_match(cols, "옵션"))
        with c3: sel_avail = st.selectbox("가용재고", cols, index=auto_match(cols, "가용재고"))
        with c4: sel_t7 = st.selectbox("7일판매량", cols, index=auto_match(cols, "7일판매량"))
        
        with st.expander("추가 매칭 항목 확인 (원가, 거래처 등)"):
            c5, c6, c7 = st.columns(3)
            with c5: st.selectbox("원가", cols, index=auto_match(cols, "원가"))
            with c6: st.selectbox("거래처", cols, index=auto_match(cols, "거래처"))
            with c7: st.selectbox("바코드", cols, index=auto_match(cols, "바코드"))

        st.divider()
        st.subheader("⚙️ 3단계: 데이터 분석 실행")
        if st.button("🚀 분석 시작 (리오더 수치 동기화)", use_container_width=True):
            # 실제 구글 시트 연동 및 계산 로직 (여기에 들어감)
            st.session_state.df_final = st.session_state.df_raw.copy() # 임시
            st.session_state.mapping = {"item": sel_item, "opt": sel_opt, "avail": sel_avail, "t7": sel_t7}
            st.rerun()

    # --- [4~6단계: 수정 및 저장] ---
    if st.session_state.df_final is not None:
        st.divider()
        st.subheader("📝 4~5단계: 수량 검토 및 수정")
        st.info("데이터 분석이 완료되었습니다. 아래 표에서 수량을 확인하세요.")
        # (기존 데이터 에디터 로직...)
        
        if st.button("🗑️ 처음부터 다시 하기", use_container_width=True):
            reset_all()

with tab2:
    st.write("준비 중입니다.")
