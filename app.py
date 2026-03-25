import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 1. 기본 설정
KST = timezone(timedelta(hours=9))
st.set_page_config(layout="wide", page_title="저스트원 재고관리")

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

# 3. 세션 상태 관리 (초기화 핵심)
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'df_final' not in st.session_state: st.session_state.df_final = None
if 'last_fn' not in st.session_state: st.session_state.last_fn = None

st.title("📦 저스트원 통합 재고 관리 시스템")
tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 관리"])

with tab1:
    # --- [1단계: 엑셀 업로드] ---
    st.subheader("📁 1단계: 엑셀 데이터 업로드")
    up_file = st.file_uploader("파일을 선택하세요", type=['xlsx', 'xls', 'csv'], key="up_key")
    
    # [수정] 파일이 바뀌면 하단 분석 데이터를 즉시 'None'으로 밀어버림
    if up_file:
        if st.session_state.last_fn != up_file.name:
            st.session_state.df_raw = None
            st.session_state.df_final = None # 이 부분이 이전 분석 표를 날리는 핵심!
            st.session_state.last_fn = up_file.name
            
            try:
                df = pd.read_csv(up_file) if up_file.name.endswith('.csv') else pd.read_excel(up_file)
                df.columns = df.columns.str.strip()
                st.session_state.df_raw = df
                st.rerun() 
            except Exception as e:
                st.error(f"파일 읽기 오류: {e}")

    # --- [2~3단계: 매핑 및 분석 설정] ---
    # df_final이 없을 때(즉, 분석 버튼 누르기 전)만 2~3단계를 보여줌
    if st.session_state.df_raw is not None and st.session_state.df_final is None:
        st.divider()
        st.subheader("🔗 2단계: 컬럼 매핑 설정")
        cols = st.session_state.df_raw.columns.tolist()
        
        c1, c2, c3, c4 = st.columns(4)
        with c1: sel_item = st.selectbox("상품명 컬럼", cols, index=cols.index(next((c for c in cols if '상품명' in c), cols[0])))
        with c2: sel_opt = st.selectbox("옵션 컬럼", cols, index=cols.index(next((c for c in cols if '옵션' in c), cols[1] if len(cols)>1 else cols[0])))
        with c3: sel_avail = st.selectbox("가용재고 컬럼", cols, index=cols.index(next((c for c in cols if '가용재고' in c), cols[2] if len(cols)>2 else cols[0])))
        with c4: sel_t7 = st.selectbox("7일판매량 컬럼", cols, index=cols.index(next((c for c in cols if '7일판매량' in c), cols[3] if len(cols)>3 else cols[0])))
        
        st.divider()
        st.subheader("⚙️ 3단계: 데이터 분석 실행")
        if st.button("🚀 분석 시작 (리오더 수치 동기화)", use_container_width=True):
            with st.spinner("구글 시트 연동 및 계산 중..."):
                df = st.session_state.df_raw.copy()
                gs_data = load_reorder_data()
                
                # 리오더 데이터 매칭
                if not gs_data.empty:
                    gs_data['상품명'] = gs_data['상품명'].astype(str).str.strip()
                    gs_data['옵션'] = gs_data['옵션'].astype(str).str.strip()
                    df['t_n'] = df[sel_item].astype(str).str.strip()
                    df['t_o'] = df[sel_opt].astype(str).str.strip()
                    df = pd.merge(df, gs_data[['상품명', '옵션', '리오더 수량']], left_on=['t_n', 't_o'], right_on=['상품명', '옵션'], how='left')
                    df['리오더 수량'] = df['리오더 수량'].fillna(0).astype(int)
                    # 중복 컬럼 정리
                    df.drop(columns=['상품명_gs', '옵션_gs', 't_n', 't_o'], inplace=True, errors='ignore')
                else:
                    df['리오더 수량'] = 0
                
                # 계산 (컬럼명 보존을 위해 딕셔너리에 매핑 정보 임시 저장)
                st.session_state.mapping = {"item": sel_item, "opt": sel_opt, "avail": sel_avail, "t7": sel_t7}
                df['일판매량'] = (pd.to_numeric(df[sel_t7], errors='coerce').fillna(0) / 7).round(1)
                df['권장발주량'] = ((df['일판매량'] * 10) - (pd.to_numeric(df[sel_avail], errors='coerce').fillna(0) + df['리오더 수량'])).clip(lower=0).astype(int)
                
                st.session_state.df_final = df
                st.rerun()

    # --- [4~6단계: 수정 및 저장] ---
    # 분석 버튼을 눌러서 df_final이 생성된 경우에만 노출
    if st.session_state.df_final is not None:
        m = st.session_state.mapping
        st.divider()
        st.subheader("📝 4~5단계: 수량 검토 및 수정")
        
        d_cols = [m["item"], m["opt"], m["avail"], '리오더 수량', '일판매량', '권장발주량']
        edited_df = st.data_editor(st.session_state.df_final[d_cols], use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("💾 6단계: 최종 발주 확정")
        if st.button("✅ 구글 시트 전송 (기록 및 업데이트)", type="primary", use_container_width=True):
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
                st.success("🎉 구글 시트 저장 완료!")
                st.balloons()
            except Exception as e:
                st.error(f"저장 오류: {e}")

    # 화면 하단 리셋 버튼
    if st.session_state.df_raw is not None:
        st.divider()
        if st.button("🗑️ 전체 데이터 초기화"):
            st.session_state.df_raw = None
            st.session_state.df_final = None
            st.session_state.last_fn = None
            st.rerun()
