import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- [1. 구글 시트 연동 함수] ---
def get_sheet():
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        spreadsheet_key = "1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg"
        return client.open_by_key(spreadsheet_key)
    except Exception as e:
        st.error(f"구글 시트 연결 실패: {e}")
        return None

def get_recent_logs():
    """입고 날짜/수량 체크용 history 데이터 로드"""
    try:
        sh = get_sheet()
        hist_sheet = sh.worksheet("history")
        data = hist_sheet.get_all_records()
        if not data: return {}
        df_h = pd.DataFrame(data)
        df_in = df_h[pd.to_numeric(df_h['리오더입고수량'], errors='coerce') > 0].copy()
        log_dict = {}
        for (name, opt), group in df_in.groupby(['상품명', '옵션']):
            recent = group.sort_values(by='저장시간', ascending=False).head(3)
            logs = [f"{str(r['저장시간'])[5:10]}({int(r['리오더입고수량'])}개)" for _, r in recent.iterrows()]
            log_dict[(str(name).strip(), str(opt).strip())] = " / ".join(logs)
        return log_dict
    except: return {}

# --- [2. 앱 설정 및 탭 정의] ---
st.set_page_config(layout="wide", page_title="재고 관리 시스템")
st.title("📦 통합 재고 관리 및 발주 시스템")

tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 및 히스토리"])

if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'in_logs' not in st.session_state: st.session_state.in_logs = {}

def get_idx(cols, keywords):
    for key in keywords:
        for i, c in enumerate(cols):
            if key in str(c): return i
    return 0

# --- [🏭 탭 1: 제작 상품 관리] ---
with tab1:
    uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'xls', 'csv'], key="main_upload")
    
    if uploaded_file is not None and st.session_state.df_raw is None:
        df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
        df.columns = df.columns.str.strip()
        df = df.loc[:, ~df.columns.duplicated()]

        # [기능 1] 리오더 수량 초기화 방지 (구글 시트 로드)
        try:
            sh = get_sheet()
            gs_data = pd.DataFrame(sh.sheet1.get_all_records())
            if not gs_data.empty and '리오더 수량' in gs_data.columns:
                p_col = df.columns[get_idx(df.columns, ['상품명'])]
                o_col = df.columns[get_idx(df.columns, ['옵션'])]
                df['match_key'] = df[p_col].astype(str).str.strip() + df[o_col].astype(str).str.strip()
                gs_data['match_key'] = gs_data['상품명'].astype(str).str.strip() + gs_data['옵션'].astype(str).str.strip()
                reorder_map = gs_data.set_index('match_key')['리오더 수량'].to_dict()
                df['입고예정수량(리오더)'] = df['match_key'].map(reorder_map).fillna(0).astype(int)
                df.drop(columns=['match_key'], inplace=True)
            else: df['입고예정수량(리오더)'] = 0
            st.session_state.in_logs = get_recent_logs()
        except: df['입고예정수량(리오더)'] = 0

        st.session_state.df_raw = df
        st.rerun()

    if st.session_state.df_raw is not None:
        cols = st.session_state.df_raw.columns.tolist()

        # [1~2단계: 매핑 설정]
        st.subheader("⚙️ 1~2단계: 자동 매핑 설정")
        c1, c2 = st.columns(2)
        sold_out = c1.selectbox("품절 여부", cols, index=get_idx(cols, ['품절', '판매중단']))
        vendor = c1.selectbox("공급처", cols, index=get_idx(cols, ['공급처', '업체명']))
        item = c1.selectbox("상품명", cols, index=get_idx(cols, ['상품명', '상품']))
        option = c1.selectbox("옵션", cols, index=get_idx(cols, ['옵션']))
        vendor_opt = c1.selectbox("공급처옵션", cols, index=get_idx(cols, ['공급처옵션', '거래처옵션']))
        reg_date = c1.selectbox("상품등록일", cols, index=get_idx(cols, ['등록일', '상품등록'])) # 추가됨
        
        stock = c2.selectbox("정상재고", cols, index=get_idx(cols, ['정상재고', '재고']))
        avail = c2.selectbox("가용재고", cols, index=get_idx(cols, ['가용재고', '가용']))
        t3day = c2.selectbox("3일 발주 합계", cols, index=get_idx(cols, ['3일', '최근3일']))
        t1week = c2.selectbox("1주 발주 합계", cols, index=get_idx(cols, ['1주', '7일', '최근7일']))

        # [3단계: 분석 및 파라미터 설정]
        st.divider()
        st.subheader("🚀 3단계: 분석 파라미터 및 실행")
        l1, l2 = st.columns(2)
        lead_time = l1.number_input("리드타임(Lead Time)", value=10)
        safety_stock = l2.number_input("안전재고(Safety Stock)", value=7)

        if st.button("📊 분석 실행"):
            df = st.session_state.df_raw
            df['일일 판매량'] = (pd.to_numeric(df[t3day], errors='coerce').fillna(0) / 3).round(1)
            df['권장 발주량'] = ((df['일일 판매량'] * (lead_time + safety_stock)) - 
                             (pd.to_numeric(df[avail], errors='coerce').fillna(0) + df['입고예정수량(리오더)'])).clip(lower=0).round(0).astype(int)
            
            # [기능 2] 품절건 발주 제외 로직
            df.loc[df[sold_out].astype(str).str.contains('품절', na=False), '권장 발주량'] = 0
            st.rerun()

        # [4단계: 실시간 입고 관리]
        st.subheader("📊 4단계: 실시간 입고 및 재고 관리")
        
        # [기능 3] 입고 날짜/수량 체크 셀 매칭
        def match_log(row):
            key = (str(row[item]).strip(), str(row[option]).strip())
            return st.session_state.in_logs.get(key, "-")
        st.session_state.df_raw['최근입고기록'] = st.session_state.df_raw.apply(match_log, axis=1)

        search = st.text_input("🔍 상품명 검색")
        df_disp = st.session_state.df_raw.copy()
        if search: df_disp = df_disp[df_disp[item].astype(str).str.contains(search, na=False)]

        if "리오더입고수량" not in df_disp.columns: df_disp["리오더입고수량"] = 0
        
        # 모든 항목 리스트 (누락 없음)
        all_cols = [sold_out, vendor, item, option, vendor_opt, reg_date, stock, avail, "입고예정수량(리오더)", "리오더입고수량", "최근입고기록", t3day, t1week, '권장 발주량']
        final_cols = [c for c in all_cols if c in df_disp.columns or c in ["리오더입고수량", "최근입고기록"]]
        
        def on_edit_main():
            changes = st.session_state["main_editor"]["edited_rows"]
            for idx_str, change in changes.items():
                idx = int(idx_str)
                orig_idx = df_disp.index[idx]
                if "리오더입고수량" in change:
                    in_qty = int(change["리오더입고수량"])
                    st.session_state.df_raw.at[orig_idx, "입고예정수량(리오더)"] -= in_qty
                    # history 기록 저장
                    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    sh = get_sheet()
                    sh.worksheet("history").append_row([now_str, str(df_disp.at[orig_idx, item]), str(df_disp.at[orig_idx, option]), in_qty])
            
            # 구글 시트 Sheet1 실시간 저장 (초기화 방지)
            sh = get_sheet()
            save_df = st.session_state.df_raw[[item, option, "입고예정수량(리오더)"]].copy()
            save_df.columns = ['상품명', '옵션', '리오더 수량']
            sh.sheet1.update('A1', [save_df.columns.values.tolist()] + save_df.values.tolist())
            st.session_state.in_logs = get_recent_logs()
            st.rerun()

        st.data_editor(df_disp[final_cols], use_container_width=True, key="main_editor", on_change=on_edit_main)

        # 5단계: 요약 및 다운로드
        if '권장 발주량' in st.session_state.df_raw.columns:
            st.subheader("📋 5단계: 최종 발주 리스트 요약")
            to_order = st.session_state.df_raw[st.session_state.df_raw['권장 발주량'] > 0]
            st.dataframe(to_order[final_cols], use_container_width=True)
            
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                to_order.to_excel(writer, index=False)
            st.download_button("📥 최종 데이터 다운로드", data=output.getvalue(), file_name=f"발주서_{datetime.now().strftime('%m%d')}.xlsx")

# --- [🌙 탭 2: 동대문 관리 및 히스토리] ---
with tab2:
    st.subheader("🌙 동대문 사입 및 전체 히스토리")
    # 히스토리 조회 버튼을 위로 배치
    if st.button("🔄 시트 기록 새로고침"):
        sh = get_sheet()
        hist_df = pd.DataFrame(sh.worksheet("history").get_all_records())
        st.dataframe(hist_df.sort_values(by='저장시간', ascending=False), use_container_width=True)
    
    st.divider()
    dong_file = st.file_uploader("사입 리스트 업로드", type=['xlsx', 'csv'], key="dong_upload")
    if dong_file:
        df_dong = pd.read_excel(dong_file) if not dong_file.name.endswith('.csv') else pd.read_csv(dong_file)
        st.data_editor(df_dong, use_container_width=True)
