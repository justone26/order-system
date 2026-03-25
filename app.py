import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

# --- [0. 상수 설정] ---
KST = timezone(timedelta(hours=9)) # 한국 시간대 설정

# --- [1. 공통 함수 정의] ---
def get_sheet():
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        return client.open_by_key("1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg")
    except Exception as e:
        st.error(f"구글 시트 연결 실패: {e}")
        return None

def save_reorder_data(df, item_col, opt_col):
    """현재 세션의 리오더 수량을 구글 시트 메인 탭에 덮어쓰기 (보존용)"""
    try:
        spreadsheet = get_sheet()
        if spreadsheet:
            sheet = spreadsheet.sheet1
            sheet.clear()
            # 상품명, 옵션, 리오더 수량 컬럼만 추출하여 저장
            save_df = df[[item_col, opt_col, '리오더 수량']].copy()
            save_df.columns = ['상품명', '옵션', '리오더 수량']
            sheet.update([save_df.columns.values.tolist()] + save_df.values.tolist())
    except Exception as e:
        st.error(f"데이터 보존 실패: {e}")

# --- [2. 앱 설정] ---
st.set_page_config(layout="wide", page_title="저스트원 재고관리")
st.title("📦 저스트원 통합 재고 관리 시스템")

tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 관리"])

with tab1:
    if 'df_raw' not in st.session_state: st.session_state.df_raw = None
    if 'analyzed' not in st.session_state: st.session_state.analyzed = False

    st.subheader("📁 1단계: 데이터 업로드")
    c_up1, c_up2 = st.columns([3, 1])
    uploaded_file = c_up1.file_uploader("엑셀/CSV 파일을 선택하세요", type=['xlsx', 'xls', 'csv'], key="main_up")
    
    if c_up2.button("🔄 전체 데이터 초기화", use_container_width=True):
        st.session_state.clear()
        st.rerun()

    # --- [핵심] 업체별 데이터 누적 및 리오더 보존 로직 ---
    if uploaded_file is not None:
        if st.session_state.get('last_fn') != uploaded_file.name:
            with st.spinner(f'{uploaded_file.name} 데이터를 처리 중입니다...'):
                # 1. 새 파일 읽기
                df_new = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
                df_new.columns = df_new.columns.str.strip()
                df_new = df_new.loc[:, ~df_new.columns.duplicated()]

                # 2. 리오더 수량 매칭 (구글 시트 연동)
                try:
                    gs = get_sheet()
                    gs_data = pd.DataFrame(gs.sheet1.get_all_records())
                    if not gs_data.empty and '리오더 수량' in gs_data.columns:
                        tmp_item = next((c for c in df_new.columns if '상품명' in c), df_new.columns[0])
                        tmp_opt = next((c for c in df_new.columns if '옵션' in c), df_new.columns[1])
                        
                        df_new['tmp_n'] = df_new[tmp_item].astype(str).str.strip()
                        df_new['tmp_o'] = df_new[tmp_opt].astype(str).str.strip()
                        gs_data['상품명'] = gs_data['상품명'].astype(str).str.strip()
                        gs_data['옵션'] = gs_data['옵션'].astype(str).str.strip()
                        
                        df_new = pd.merge(df_new, gs_data[['상품명', '옵션', '리오더 수량']], 
                                         left_on=['tmp_n', 'tmp_o'], right_on=['상품명', '옵션'], 
                                         how='left', suffixes=('', '_gs'))
                        
                        if '리오더 수량_gs' in df_new.columns:
                            df_new['리오더 수량'] = df_new['리오더 수량_gs'].fillna(0).astype(int)
                            df_new.drop(columns=['상품명_gs', '옵션_gs', '리오더 수량_gs', 'tmp_n', 'tmp_o'], inplace=True, errors='ignore')
                except:
                    if '리오더 수량' not in df_new.columns: df_new['리오더 수량'] = 0

                # 3. [데이터 누적] 기존 데이터가 있으면 합치기
                if st.session_state.df_raw is not None:
                    # 기존 데이터 + 새 데이터 합치기
                    st.session_state.df_raw = pd.concat([st.session_state.df_raw, df_new], ignore_index=True)
                    # 중복 제거 (상품명/옵션 기준)
                    target_n = next((c for c in st.session_state.df_raw.columns if '상품명' in c), st.session_state.df_raw.columns[0])
                    target_o = next((c for c in st.session_state.df_raw.columns if '옵션' in c), st.session_state.df_raw.columns[1])
                    st.session_state.df_raw.drop_duplicates(subset=[target_n, target_o], keep='last', inplace=True)
                    st.success(f"✅ 기존 데이터에 {uploaded_file.name} 업체가 추가되었습니다!")
                else:
                    st.session_state.df_raw = df_new
                    st.success(f"✅ {uploaded_file.name} 데이터가 로드되었습니다.")

                st.session_state.last_fn = uploaded_file.name
                st.session_state.analyzed = False
                st.rerun()

    # --- [매핑 및 분석 로직] ---
    if st.session_state.df_raw is not None:
        df_curr = st.session_state.df_raw
        cols = df_curr.columns.tolist()

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

        if st.button("🚀 분석 실행", use_container_width=True, type="primary"):
            st.session_state.analyzed = True
            st.rerun()

        if st.session_state.analyzed:
            st.divider()
            # 여기에 사장님의 4~6단계 로직(계산 및 편집)을 붙여넣으시면 됩니다.
            # 시간 기록 시 datetime.now(KST)를 사용하면 한국 시간으로 저장됩니다.
            st.info(f"현재 {len(st.session_state.df_raw)}개의 상품 데이터가 통합되어 있습니다.")
