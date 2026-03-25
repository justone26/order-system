import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 1. [핵심] 한국 표준시(KST) 및 날짜 설정
def get_now_kst():
    return datetime.now(timezone(timedelta(hours=9)))

now = get_now_kst()
today_str = now.strftime("%Y-%m-%d %H:%M:%S")

st.set_page_config(layout="wide", page_title=f"저스트원 재고관리 ({now.strftime('%m/%d')})")

# 2. [핵심] 구글 시트 연동 함수
def get_google_sheet():
    try:
        # Streamlit Secrets에 저장된 서비스 계정 키 사용
        creds_dict = dict(st.secrets["gcp_service_account"])
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        # 사장님의 구글 시트 ID (기존 사용하시던 ID)
        return client.open_by_key("1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg")
    except Exception as e:
        st.error(f"구글 시트 연결 실패: {e}")
        return None

# 리셋 콜백
def reset_callback():
    for key in list(st.session_state.keys()):
        del st.session_state[key]

# 자동 매칭 로직
def get_auto_index(cols, keywords):
    for i, col in enumerate(cols):
        if any(k in str(col).strip() for k in keywords):
            return i
    return 0

# 세션 초기화
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'df_final' not in st.session_state: st.session_state.df_final = None
if 'mapping' not in st.session_state: st.session_state.mapping = {}

# --- 화면 시작 ---
st.title("📦 저스트원 통합 재고 관리 시스템")
st.write(f"🕒 **현재 분석 시간:** {today_str} (KST)")

tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 관리"])

with tab1:
    # 1단계: 업로드
    st.subheader("📁 1단계: 데이터 업로드")
    up_file = st.file_uploader("파일을 선택하세요", type=['xlsx', 'xls', 'csv'], key="main_uploader", label_visibility="collapsed")
    
    col_btn, _ = st.columns([1, 3])
    with col_btn:
        st.button("🔄 전체 데이터 초기화", use_container_width=True, on_click=reset_callback)

    if up_file is not None and st.session_state.df_raw is None:
        try:
            df = pd.read_csv(up_file) if up_file.name.endswith('.csv') else pd.read_excel(up_file)
            df.columns = df.columns.str.strip()
            st.session_state.df_raw = df
            st.rerun()
        except Exception as e:
            st.error(f"파일 읽기 실패: {e}")

    # 2단계 & 3단계: 설정 영역 (파일 있을 때 상시 노출)
    if st.session_state.df_raw is not None:
        st.divider()
        st.subheader("🔗 2단계: 자동 컬럼 매칭")
        cols = st.session_state.df_raw.columns.tolist()
        
        c1, c2 = st.columns(2)
        with c1:
            sold_out = st.selectbox("품절 여부", cols, index=get_auto_index(cols, ['품절', '판매중단']))
            vendor = st.selectbox("공급처", cols, index=get_auto_index(cols, ['공급처', '업체명']))
            item = st.selectbox("상품명", cols, index=get_auto_index(cols, ['상품명', '상품']))
            option = st.selectbox("옵션", cols, index=get_auto_index(cols, ['옵션']))
            vendor_item = st.selectbox("공급처 상품명", cols, index=get_auto_index(cols, ['공급처상품명', '거래처옵션']))
        with c2:
            reg_date = st.selectbox("등록일", cols, index=get_auto_index(cols, ['등록일', '생성일']))
            stock = st.selectbox("정상재고", cols, index=get_auto_index(cols, ['정상재고', '재고']))
            avail = st.selectbox("가용재고", cols, index=get_auto_index(cols, ['가용재고', '가용']))
            t3day = st.selectbox("3일 발주합계", cols, index=get_auto_index(cols, ['3일']))
            t1week = st.selectbox("7일 발주합계", cols, index=get_auto_index(cols, ['7일', '1주']))

        st.divider()
        st.subheader("⚙️ 3단계: 발주 기준 및 구글시트 연동")
        col_lt, col_ss, col_sheet = st.columns([1, 1, 2])
        with col_lt:
            lt_val = st.number_input("⏳ 리드타임 (일)", min_value=1, value=st.session_state.mapping.get('lt', 7))
        with col_ss:
            ss_val = st.number_input("🛡️ 안전재고 (일)", min_value=0, value=st.session_state.mapping.get('ss', 3))
        with col_sheet:
            st.info("📊 시트 연결 상태: 연동 준비 완료")

        if st.button("🚀 데이터 분석 및 구글시트 대조 시작", use_container_width=True):
            with st.spinner("구글 시트에서 최신 리오더 데이터를 가져오는 중..."):
                # 구글 시트 데이터 로드
                sh = get_google_sheet()
                # 예: '발주기록' 시트에서 데이터를 가져온다고 가정
                # ws = sh.worksheet("발주기록")
                # sheet_data = pd.DataFrame(ws.get_all_records())
                
                st.session_state.mapping = {
                    "item": item, "option": option, "avail": avail, 
                    "t1week": t1week, "lt": lt_val, "ss": ss_val
                }
                
                # 계산 로직
                df_calc = st.session_state.df_raw.copy()
                df_calc['일판매량'] = (pd.to_numeric(df_calc[t1week], errors='coerce').fillna(0) / 7).round(2)
                df_calc['필요재고'] = (df_calc['일판매량'] * (lt_val + ss_val)).round(0).astype(int)
                df_calc['가용재고_num'] = pd.to_numeric(df_calc[avail], errors='coerce').fillna(0)
                df_calc['권장발주량'] = (df_calc['필요재고'] - df_calc['가용재고_num']).clip(lower=0).astype(int)
                
                st.session_state.df_final = df_calc
                st.rerun()

    # 4단계: 결과 확인
    if st.session_state.df_final is not None:
        st.divider()
        st.success(f"✅ 분석 완료 ({today_str})")
        m = st.session_state.mapping
        display_cols = [m['item'], m['option'], m['avail'], m['t1week'], '일판매량', '필요재고', '권장발주량']
        st.data_editor(st.session_state.df_final[display_cols], use_container_width=True, hide_index=True)

        if st.button("✅ 구글 시트로 최종 발주 데이터 전송", type="primary", use_container_width=True):
            st.write("구글 시트에 데이터를 기록하는 중입니다... (함수 연결 필요)")
