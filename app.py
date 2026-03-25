import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 1. 기본 설정 (KST 시간 및 페이지 넓게 쓰기)
KST = timezone(timedelta(hours=9))
st.set_page_config(layout="wide", page_title="저스트원 재고관리")

# 2. 구글 시트 연결 및 데이터 처리 함수
def get_sheet():
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        # 사장님 구글 시트 ID
        return client.open_by_key("1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg")
    except:
        return None

def load_reorder_data():
    """Sheet1(리오더데이터)에서 기존 수치 로드"""
    try:
        ss = get_sheet()
        if ss:
            return pd.DataFrame(ss.sheet1.get_all_records())
        return pd.DataFrame()
    except:
        return pd.DataFrame()

# 3. 세션 상태 초기화 (단계별 진행을 위한 톱니바퀴)
if 'step' not in st.session_state: st.session_state.step = 1
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'df_final' not in st.session_state: st.session_state.df_final = None

st.title("📦 저스트원 통합 재고 관리 시스템")
tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 관리"])

with tab1:
    # --- [1단계: 엑셀 업로드] ---
    st.subheader("📁 1단계: 엑셀 데이터 업로드")
    up_file = st.file_uploader("파일을 선택하세요 (xlsx, csv)", type=['xlsx', 'xls', 'csv'], key="up_key")
    
    if up_file and st.session_state.df_raw is None:
        try:
            df = pd.read_csv(up_file) if up_file.name.endswith('.csv') else pd.read_excel(up_file)
            df.columns = df.columns.str.strip()
            st.session_state.df_raw = df
            st.session_state.step = 2  # 업로드 성공하면 2단계로
            st.rerun()
        except Exception as e:
            st.error(f"파일 읽기 오류: {e}")

    # --- [2단계: 컬럼 매핑 설정] ---
    if st.session_state.step >= 2 and st.session_state.df_raw is not None:
        st.divider()
        st.subheader("🔗 2단계: 컬럼 매핑 설정")
        cols = st.session_state.df_raw.columns.tolist()
        
        c1, c2, c3, c4 = st.columns(4)
        with c1: sel_item = st.selectbox("상품명 컬럼", cols, index=cols.index(next((c for c in cols if '상품명' in c), cols[0])))
        with c2: sel_opt = st.selectbox("옵션 컬럼", cols, index=cols.index(next((c for c in cols if '옵션' in c), cols[1] if len(cols)>1 else cols[0])))
        with c3: sel_avail = st.selectbox("가용재고 컬럼", cols, index=cols.index(next((c for c in cols if '가용재고' in c), cols[2] if len(cols)>2 else cols[0])))
        with c4: sel_t7 = st.selectbox("7일판매량 컬럼", cols, index=cols.index(next((c for c in cols if '7일판매량' in c), cols[3] if len(cols)>3 else cols[0])))
        
        # --- [3단계: 분석 실행 버튼] ---
        st.divider()
        st.subheader("⚙️ 3단계: 데이터 분석 및 리오더 동기화")
        if st.button("🚀 분석 시작 (구글 시트 데이터 불러오기)", use_container_width=True):
            with st.spinner("구글 시트에서 기존 리오더 수치를 가져오는 중..."):
                df = st.session_state.df_raw.copy()
                # 구글 시트 데이터 합치기
                gs_data = load_reorder_data()
                if not gs_data.empty:
                    gs_data['상품명'] = gs_data['상품명'].astype(str).str.strip()
                    gs_data['옵션'] = gs_data['옵션'].astype(str).str.strip()
                    df['t_n'] = df[sel_item].astype(str).str.strip()
                    df['t_o'] = df[sel_opt].astype(str).str.strip()
                    
                    df = pd.merge(df, gs_data[['상품명', '옵션', '리오더 수량']], 
                                  left_on=['t_n', 't_o'], right_on=['상품명', '옵션'], 
                                  how='left', suffixes=('', '_gs'))
                    df['리오더 수량'] = df['리오더 수량'].fillna(0).astype(int)
                    df.drop(columns=['상품명_gs', '옵션_gs', 't_n', 't_o'], inplace=True, errors='ignore')
                else:
                    df['리오더 수량'] = 0
                
                # 일판매량 및 권장발주량 자동 계산
                df['일판매량'] = (pd.to_numeric(df[sel_t7], errors='coerce').fillna(0) / 7).round(1)
                df['권장발주량'] = ((df['일판매량'] * 10) - (pd.to_numeric(df[sel_avail], errors='coerce').fillna(0) + df['리오더 수량'])).clip(lower=0).astype(int)
                
                st.session_state.df_final = df
                st.session_state.step = 4
                st.rerun()

    # --- [4~5단계: 최종 수정 및 6단계: 저장] ---
    if st.session_state.step >= 4 and st.session_state.df_final is not None:
        st.divider()
        st.subheader("📝 4~5단계: 최종 수량 검토 및 수정")
        st.info("💡 표 안의 숫자를 더블클릭해서 수정할 수 있습니다.")
        
        d_cols = [sel_item, sel_opt, sel_avail, '리오더 수량', '일판매량', '권장발주량']
        edited_df = st.data_editor(st.session_state.df_final[d_cols], use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("💾 6단계: 구글 시트 전송 및 발주 확정")
        if st.button("✅ 최종 저장 (발주기록 작성 및 리오더 업데이트)", type="primary", use_container_width=True):
            try:
                ss = get_sheet()
                # 1. '발주기록' 탭에 누적 저장
                log_ws = ss.worksheet("발주기록")
                now_str = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
                to_log = edited_df[edited_df['권장발주량'] > 0]
                
                if not to_log.empty:
                    rows = [[now_str, r[sel_item], r[sel_opt], int(r[sel_avail]), int(r['권장발주량'])] for _, r in to_log.iterrows()]
                    log_ws.append_rows(rows)
                
                # 2. 'Sheet1'에 현재 리오더 상태 갱신 (다음 로딩을 위해)
                sh1 = ss.sheet1
                sh1.clear()
                # 제목줄 + 데이터
                header = [['상품명', '옵션', '리오더 수량']]
                data = edited_df[[sel_item, sel_opt, '리오더 수량']].values.tolist()
                sh1.update(header + data)
                
                st.success("🎉 구글 시트에 안전하게 기록되었습니다! 리오더 수치 보존 완료.")
                st.balloons()
            except Exception as e:
                st.error(f"저장 중 오류 발생 (시트 탭 이름을 확인하세요): {e}")

    # 리셋 (처음으로 돌아가기)
    if st.session_state.step > 1:
        st.divider()
        if st.button("🗑️ 화면 데이터 싹 비우기 (처음부터 다시)"):
            st.session_state.step = 1
            st.session_state.df_raw = None
            st.session_state.df_final = None
            st.rerun()

with tab2:
    st.write("🌙 동대문 사입 관리 기능은 준비 중입니다.")
