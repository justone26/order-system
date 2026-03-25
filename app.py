import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

# --- [0. 상수 설정] ---
KST = timezone(timedelta(hours=9))

# --- [1. 공통 함수 정의] ---
def get_sheet():
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        # 사장님 시트 ID 고정
        return client.open_by_key("1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg")
    except Exception as e:
        st.error(f"구글 시트 연결 실패: {e}")
        return None

def save_reorder_data(df):
    """현재 세션의 리오더 수량을 구글 시트 메인 탭에 덮어쓰기 (보존용)"""
    try:
        spreadsheet = get_sheet()
        if spreadsheet:
            sheet = spreadsheet.sheet1
            sheet.clear()
            # 상품명, 옵션, 리오더 수량만 저장하여 데이터 용량 최적화
            save_df = df[['상품명', '옵션', '리오더 수량']].copy()
            sheet.update([save_df.columns.values.tolist()] + save_df.values.tolist())
    except Exception as e:
        st.error(f"데이터 보존 실패: {e}")

# --- [2. 앱 설정] ---
st.set_page_config(layout="wide", page_title="저스트원 재고관리")
st.title("📦 저스트원 통합 재고 관리 시스템")

tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 관리"])

with tab1:
    # 세션 상태 초기화
    if 'df_raw' not in st.session_state: st.session_state.df_raw = None
    if 'analyzed' not in st.session_state: st.session_state.analyzed = False

    st.subheader("📁 1단계: 데이터 업로드")
    uploaded_file = st.file_uploader("엑셀/CSV 파일을 선택하세요", type=['xlsx', 'xls', 'csv'])

    # [핵심] 파일 업로드 시 리오더 수량 보존 로직
    if uploaded_file is not None:
        # 파일이 처음 올라오거나 다른 파일이 올라왔을 때만 실행
        if st.session_state.get('last_fn') != uploaded_file.name:
            with st.spinner('데이터를 불러오며 기존 리오더 수량을 매칭 중입니다...'):
                # 1. 새 파일 읽기
                df_new = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
                df_new.columns = df_new.columns.str.strip()
                
                # 2. 구글 시트에서 어제 저장한 '리오더 수량' 불러오기
                try:
                    gs = get_sheet()
                    gs_data = pd.DataFrame(gs.sheet1.get_all_records())
                    
                    if not gs_data.empty and '리오더 수량' in gs_data.columns:
                        # 상품명과 옵션을 기준으로 기존 수량 합치기
                        # 임시 매핑 (파일의 상품명/옵션 컬럼 자동 찾기)
                        tmp_item = next((c for c in df_new.columns if '상품명' in c), df_new.columns[0])
                        tmp_opt = next((c for c in df_new.columns if '옵션' in c), df_new.columns[1])
                        
                        df_new['tmp_n'] = df_new[tmp_item].astype(str).str.strip()
                        df_new['tmp_o'] = df_new[tmp_opt].astype(str).str.strip()
                        gs_data['상품명'] = gs_data['상품명'].astype(str).str.strip()
                        gs_data['옵션'] = gs_data['옵션'].astype(str).str.strip()
                        
                        # 병합(Merge)
                        df_new = pd.merge(df_new, gs_data[['상품명', '옵션', '리오더 수량']], 
                                         left_on=['tmp_n', 'tmp_o'], right_on=['상품명', '옵션'], 
                                         how='left', suffixes=('', '_old'))
                        
                        # 기존 수량이 있으면 쓰고, 없으면 0
                        if '리오더 수량_old' in df_new.columns:
                            df_new['리오더 수량'] = df_new['리오더 수량_old'].fillna(0).astype(int)
                            df_new.drop(columns=['상품명_old', '옵션_old', '리오더 수량_old', 'tmp_n', 'tmp_o'], inplace=True, errors='ignore')
                except:
                    if '리오더 수량' not in df_new.columns: df_new['리오더 수량'] = 0
                
                st.session_state.df_raw = df_new
                st.session_state.last_fn = uploaded_file.name
                st.session_state.analyzed = False
                st.rerun()

    # --- [매핑 및 분석 로직] ---
    if st.session_state.df_raw is not None:
        df_curr = st.session_state.df_raw
        cols = df_curr.columns.tolist()

        # 사장님이 주신 10개 매핑 항목
        st.divider()
        st.subheader("⚙️ 2단계: 매핑 설정")
        c1, c2 = st.columns(2)
        with c1:
            sold_out = st.selectbox("품절 여부", cols, index=0)
            vendor = st.selectbox("공급처", cols, index=0)
            v_item = st.selectbox("공급처 상품명", cols, index=0)
            item = st.selectbox("상품명", cols, index=0)
            option = st.selectbox("옵션", cols, index=0)
        with c2:
            reg_date = st.selectbox("등록일", cols, index=0)
            stock = st.selectbox("정상재고", cols, index=0)
            avail = st.selectbox("가용재고", cols, index=0)
            t3day = st.selectbox("3일 발주합계", cols, index=0)
            t7day = st.selectbox("7일 발주합계", cols, index=0)

        # 분석 실행 버튼
        if st.button("🚀 분석 실행", use_container_width=True, type="primary"):
            st.session_state.analyzed = True
            st.rerun()

        # --- [결과 화면] ---
        if st.session_state.analyzed:
            st.divider()
            # (이하 생략 - 사장님 소스의 4~6단계 로직을 여기에 배치)
            st.info("매핑이 완료되었습니다. 이제 아래에서 리오더 수량을 편집하고 저장하면 구글 시트에 즉시 반영됩니다.")
