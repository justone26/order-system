import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 1. 기본 설정
KST = timezone(timedelta(hours=9))
st.set_page_config(layout="wide", page_title="저스트원 재고관리")

# [보강] 사장님이 말씀하신 10가지 자동 매칭 키워드 리스트
MATCH_KEYS = {
    "상품명": ["상품명", "물명", "아이템", "Item Name", "상품"],
    "옵션": ["옵션", "규격", "사이즈", "컬러", "Option"],
    "가용재고": ["가용재고", "현재고", "재고수량", "재고", "Stock"],
    "7일판매량": ["7일판매량", "주간판매", "7일판매", "판매량(7일)", "Sale7"],
    "원가": ["원가", "매입가", "구매가", "Cost"],
    "판매가": ["판매가", "매출가", "정가", "Price"],
    "바코드": ["바코드", "품번", "관리코드", "Barcode"],
    "거래처": ["거래처", "제조사", "공급처", "Vendor"],
    "상태": ["상태", "판매상태", "진열상태", "Status"],
    "카테고리": ["카테고리", "분류", "Category"]
}

def auto_match(cols, target_key):
    """엑셀 컬럼 중 매칭되는 키워드가 있으면 인덱스 반환"""
    for i, col in enumerate(cols):
        if any(k in col for k in MATCH_KEYS[target_key]):
            return i
    return 0

# 2. 구글 시트 연결 함수
def get_sheet():
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        return client.open_by_key("1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg")
    except: return None

def load_reorder_data():
    try:
        ss = get_sheet()
        return pd.DataFrame(ss.sheet1.get_all_records()) if ss else pd.DataFrame()
    except: return pd.DataFrame()

# 3. 세션 상태 관리
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'df_final' not in st.session_state: st.session_state.df_final = None
if 'last_fn' not in st.session_state: st.session_state.last_fn = None

st.title("📦 저스트원 통합 재고 관리 시스템")
tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 관리"])

with tab1:
    # --- [1단계: 엑셀 업로드] ---
    st.subheader("📁 1단계: 엑셀 데이터 업로드")
    up_file = st.file_uploader("파일을 선택하세요", type=['xlsx', 'xls', 'csv'], key="up_key")
    
    if up_file and st.session_state.last_fn != up_file.name:
        st.session_state.df_raw = None
        st.session_state.df_final = None
        st.session_state.last_fn = up_file.name
        try:
            df = pd.read_csv(up_file) if up_file.name.endswith('.csv') else pd.read_excel(up_file)
            df.columns = df.columns.str.strip()
            st.session_state.df_raw = df
            st.rerun()
        except Exception as e:
            st.error(f"파일 읽기 오류: {e}")

    # --- [2~3단계: 자동 매칭 및 분석 설정] ---
    if st.session_state.df_raw is not None and st.session_state.df_final is None:
        st.divider()
        st.subheader("🔗 2단계: 자동 컬럼 매칭 (10종)")
        cols = st.session_state.df_raw.columns.tolist()
        
        # 자동 매칭 적용된 선택창
        c1, c2, c3, c4 = st.columns(4)
        with c1: sel_item = st.selectbox("상품명", cols, index=auto_match(cols, "상품명"))
        with c2: sel_opt = st.selectbox("옵션", cols, index=auto_match(cols, "옵션"))
        with c3: sel_avail = st.selectbox("가용재고", cols, index=auto_match(cols, "가용재고"))
        with c4: sel_t7 = st.selectbox("7일판매량", cols, index=auto_match(cols, "7일판매량"))
        
        # 나머지 6가지 (필요시 확장 가능하도록 UI만 배치 가능)
        with st.expander("추가 매칭 항목 확인 (원가, 거래처 등)"):
            c5, c6, c7 = st.columns(3)
            with c5: st.selectbox("원가", cols, index=auto_match(cols, "원가"))
            with c6: st.selectbox("거래처", cols, index=auto_match(cols, "거래처"))
            with c7: st.selectbox("바코드", cols, index=auto_match(cols, "바코드"))

        st.divider()
        st.subheader("⚙️ 3단계: 분석 실행")
        if st.button("🚀 데이터 분석 및 리오더 동기화", use_container_width=True):
            with st.spinner("구글 시트 연동 중..."):
                df = st.session_state.df_raw.copy()
                gs_data = load_reorder_data()
                
                if not gs_data.empty:
                    gs_data['상품명'] = gs_data['상품명'].astype(str).str.strip()
                    gs_data['옵션'] = gs_data['옵션'].astype(str).str.strip()
                    df['t_n'] = df[sel_item].astype(str).str.strip()
                    df['t_o'] = df[sel_opt].astype(str).str.strip()
                    df = pd.merge(df, gs_data[['상품명', '옵션', '리오더 수량']], left_on=['t_n', 't_o'], right_on=['상품명', '옵션'], how='left')
                    df['리오더 수량'] = df['리오더 수량'].fillna(0).astype(int)
                    df.drop(columns=['상품명_gs', '옵션_gs', 't_n', 't_o'], inplace=True, errors='ignore')
                else:
                    df['리오더 수량'] = 0
                
                st.session_state.mapping = {"item": sel_item, "opt": sel_opt, "avail": sel_avail, "t7": sel_t7}
                df['일판매량'] = (pd.to_numeric(df[sel_t7], errors='coerce').fillna(0) / 7).round(1)
                df['권장발주량'] = ((df['일판매량'] * 10) - (pd.to_numeric(df[sel_avail], errors='coerce').fillna(0) + df['리오더 수량'])).clip(lower=0).astype(int)
                
                st.session_state.df_final = df
                st.rerun()

    # --- [4~6단계: 수정 및 저장] ---
    if st.session_state.df_final is not None:
        m = st.session_state.mapping
        st.divider()
        st.subheader("📝 4~5단계: 수량 검토 및 수정")
        d_cols = [m["item"], m["opt"], m["avail"], '리오더 수량', '일판매량', '권장발주량']
        edited_df = st.data_editor(st.session_state.df_final[d_cols], use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("💾 6단계: 최종 발주 확정")
        if st.button("✅ 구글 시트 저장 및 기록", type="primary", use_container_width=True):
            try:
                ss = get_sheet()
                log_ws = ss.worksheet("발주기록")
                now_str = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
                to_log = edited_df[edited_df['권장발주량'] > 0]
                if not to_log.empty:
                    rows = [[now_str, r[m["item"]], r[m["opt"]], int(r[m["avail"]]), int(r['권장발주량'])] for _, r in to_log.iterrows()]
                    log_ws.append_rows(rows)
                
                sh1 = ss.sheet1
                sh1.clear()
                sh1.update([['상품명', '옵션', '리오더 수량']] + edited_df[[m["item"], m["opt"], '리오더 수량']].values.tolist())
                st.success("✅ 저장 완료!")
                st.balloons()
            except Exception as e:
                st.error(f"저장 오류: {e}")

    # 리셋
    if st.session_state.df_raw is not None:
        if st.button("🗑️ 전체 초기화"):
            st.session_state.df_raw = None
            st.session_state.df_final = None
            st.session_state.last_fn = None
            st.rerun()
