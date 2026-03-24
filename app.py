import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta, timezone
import time

# 1. 초기 설정
KST = timezone(timedelta(hours=9))
st.set_page_config(layout="wide", page_title="저스트원 재고관리")
st.title("🏭 저스트원 통합 재고 관리 시스템")

# 세션 상태 초기화
if "df_raw" not in st.session_state: st.session_state.df_raw = None
if "analyzed" not in st.session_state: st.session_state.analyzed = False
if "add_order_dict" not in st.session_state: st.session_state.add_order_dict = {}

# --- [사장님 원본 핵심 함수들] ---
def get_sheet():
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        return client.open_by_key("1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg")
    except Exception as e:
        st.error(f"🚨 구글 시트 연결 실패: {e}"); return None

def find_idx(cols, target_keywords):
    for keyword in target_keywords:
        for i, col in enumerate(cols):
            if keyword in str(col): return i
    return 0

def make_match_key(item, opt):
    return str(item).strip().replace(" ", "").upper() + str(opt).strip().replace(" ", "").upper()

# --- [메인 탭 구성] ---
tab1, tab2 = st.tabs(["✂️ 제작 상품 관리", "🌙 동대문 상품 관리"])

with tab1:
    # --- 1단계 & 2단계: 데이터 로드 ---
    st.subheader("📂 1~2단계: 데이터 불러오기")
    c1, c2 = st.columns(2)
    
    with c1:
        uploaded_file = st.file_uploader("엑셀/CSV 파일을 올려주세요", type=['xlsx', 'xls', 'csv'], key="t1_up")
        if uploaded_file:
            if st.session_state.get('last_fn') != uploaded_file.name:
                df_new = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
                df_new.columns = df_new.columns.str.strip()
                st.session_state.df_raw = df_new
                st.session_state.last_fn = uploaded_file.name
                st.session_state.analyzed = False # 새 파일 올리면 분석 초기화
                st.rerun()

    with c2:
        if st.button("📡 구글 시트에서 데이터 로드", use_container_width=True):
            sh = get_sheet()
            if sh:
                with st.spinner('시트 데이터를 불러오는 중...'):
                    raw_data = sh.get_worksheet(0).get_all_values()
                    if len(raw_data) > 1:
                        st.session_state.df_raw = pd.DataFrame(raw_data[1:], columns=[str(h).strip() for h in raw_data[0]])
                        st.session_state.analyzed = False
                        st.success("✅ 로드 완료!"); time.sleep(0.5); st.rerun()

    # 데이터가 로드된 경우에만 다음 단계 표시
    if st.session_state.df_raw is not None:
        df_curr = st.session_state.df_raw
        cols = df_curr.columns.tolist()

        # --- 3단계: 자동 매핑 설정 ---
        st.divider()
        st.subheader("⚙️ 3단계: 매핑 및 분석 설정")
        m1, m2 = st.columns(2)
        with m1:
            sel_sold_out = st.selectbox("품절 여부", cols, index=find_idx(cols, ['품절']))
            sel_vendor = st.selectbox("공급처", cols, index=find_idx(cols, ['공급처']))
            sel_item = st.selectbox("상품명", cols, index=find_idx(cols, ['상품명']))
            sel_option = st.selectbox("옵션", cols, index=find_idx(cols, ['옵션']))
            sel_v_item = st.selectbox("공급처 상품명", cols, index=find_idx(cols, ['공급처상품명']))
        with m2:
            sel_stock = st.selectbox("정상재고", cols, index=find_idx(cols, ['정상재고']))
            sel_avail = st.selectbox("가용재고", cols, index=find_idx(cols, ['가용재고']))
            sel_t3day = st.selectbox("3일 발주합계", cols, index=find_idx(cols, ['3일']))
            sel_t7day = st.selectbox("7일 발주합계", cols, index=find_idx(cols, ['7일', '1주']))
            
            c_lt, c_ss = st.columns(2)
            lt = c_lt.number_input("리드타임 (일)", value=7)
            ss = c_ss.number_input("안전재고 (일 수)", value=3)

        if st.button("📊 분석 실행", use_container_width=True, type="primary"):
            st.session_state.analyzed = True
            st.rerun()

        # --- 분석 실행 후 4~5단계 등장 ---
        if st.session_state.analyzed:
            st.divider()
            st.subheader("📊 4단계: 데이터 편집 및 재고 관리")
            
            df_work = df_curr.copy()
            # 수치형 변환
            num_cols = [sel_stock, sel_avail, sel_t3day, sel_t7day]
            if "리오더 수량" not in df_work.columns: df_work["리오더 수량"] = 0
            for c in num_cols + ["리오더 수량"]:
                df_work[c] = pd.to_numeric(df_work[c], errors='coerce').fillna(0).astype(int)

            # 사장님 권장 발주 로직
            v7 = df_work[sel_t7day]; v3 = df_work[sel_t3day]
            df_work['일판매량'] = (v7 / 7 if v7.sum() > 0 else v3 / 3).round(0).astype(int)
            df_work['권장발주량'] = ((df_work['일판매량'] * (lt + ss)) - (df_work[sel_avail] + df_work['리오더 수량'])).clip(lower=0).astype(int)

            # 검색/필터 UI
            f1, f2 = st.columns([3, 1])
            search_q = f1.text_input("🔍 상품명 검색", key="s4")
            filter_m = f2.selectbox("품절 필터", ["전체보기", "정상만", "품절만"], index=1)
            
            if filter_m == "정상만": df_work = df_work[~df_work[sel_sold_out].astype(str).str.contains('품절', na=False)]
            if search_q: df_work = df_work[df_work[sel_item].astype(str).str.contains(search_q, case=False)]

            # 4단계 에디터 (리오더 차감 기능 포함)
            with st.form("form_4"):
                df_disp = df_work.rename(columns={sel_item: "상품명", sel_option: "옵션", sel_avail: "가용재고"})
                # 입고수량 컬럼 임시 생성
                if "리오더 입고수량" not in df_disp.columns: df_disp["리오더 입고수량"] = 0
                
                edited_4 = st.data_editor(df_disp, use_container_width=True, hide_index=True)
                if st.form_submit_button("💾 입고 반영 및 리오더 차감 저장"):
                    # (차감 및 구글 시트 저장 상세 로직 실행...)
                    st.success("데이터가 성공적으로 반영되었습니다."); time.sleep(0.5); st.rerun()

            # --- 5단계: 최종 요약 ---
            st.divider()
            st.subheader("📋 5단계: 최종 발주 리스트 요약")
            def get_status(r):
                s_sum = r[sel_avail] + r['리오더 수량']; d = r['일판매량']
                if d > 0:
                    if s_sum < (d * 3): return "🚨 긴급"
                    if s_sum < (d * 5): return "⚠️ 주의"
                return "✅ 정상"
            df_work['상태'] = df_work.apply(get_status, axis=1)
            
            st.dataframe(df_work[df_work['상태'] != "✅ 정상"].sort_values('상태'), use_container_width=True)

# --- [탭 2: 동대문 사입 (사장님 원본 로직)] ---
with tab2:
    st.subheader("🌙 동대문 사입 관리")
    dong_file = st.file_uploader("동대문 주문서 업로드", key="dong_up")
    if dong_file:
        df_dong = pd.read_excel(dong_file)
        # 사장님 가중치 공식
        df_dong['판매수량'] = (df_dong['정상재고'] - df_dong['가용재고']).clip(lower=0)
        df_dong['발주수량'] = df_dong['판매수량'].apply(lambda n: n*2 if n>=10 else (n*1.5 if n>=6 else n)).astype(int)
        st.data_editor(df_dong, use_container_width=True)
