import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 1. [핵심] 한국 표준시(KST) 및 날짜 설정
def get_now_kst():
    return datetime.now(timezone(timedelta(hours=9)))

now = get_now_kst()
today_date = now.strftime("%Y-%m-%d")
today_time = now.strftime("%H:%M:%S")

st.set_page_config(layout="wide", page_title=f"저스트원 재고관리 ({now.strftime('%m/%d')})")

# [나중에 쓸 전송용 함수]
def save_to_google(df_to_save):
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        sh = client.open_by_key("1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg")
        # 여기에 저장할 시트 이름(예: '발주현황')을 넣으면 됩니다.
        ws = sh.worksheet("Sheet1") 
        # 데이터프레임을 리스트로 변환하여 추가
        ws.append_rows(df_to_save.values.tolist())
        return True
    except:
        return False

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
st.info(f"📅 **분석 기준일:** {today_date} / **현재 시간:** {today_time}")

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
        st.subheader("⚙️ 3단계: 발주 기준 설정")
        col_lt, col_ss = st.columns(2)
        with col_lt:
            lt_val = st.number_input("⏳ 리드타임 (일) - 디폴트 7일", min_value=1, value=7)
        with col_ss:
            ss_val = st.number_input("🛡️ 안전재고 (일) - 디폴트 3일", min_value=0, value=3)

        if st.button("🚀 데이터 분석 시작", use_container_width=True):
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

    # 4단계: 결과 확인 및 저장
    if st.session_state.df_final is not None:
        st.divider()
        st.subheader("📊 4~5단계: 발주 수량 검토 및 수정")
        m = st.session_state.mapping
        
        # 사장님이 편집할 결과 표 구성
        display_cols = [m['item'], m['option'], m['avail'], m['t1week'], '일판매량', '필요재고', '권장발주량']
        edited_df = st.data_editor(st.session_state.df_final[display_cols], use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("💾 6단계: 최종 데이터 저장")
        if st.button("✅ 분석된 데이터를 구글 시트로 저장하기", type="primary", use_container_width=True):
            # [연동 로직] 실제 저장은 여기서 일어납니다.
            success = save_to_google(edited_df)
            if success:
                st.success(f"🎉 성공적으로 구글 시트에 저장되었습니다! ({today_date} {today_time})")
            else:
                st.error("❌ 저장 실패! 구글 시트 권한이나 인터넷 연결을 확인해 주세요.")
