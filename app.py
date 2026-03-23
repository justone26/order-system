import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- [1. 공통 함수 및 구글 시트 로직] ---
def get_sheet():
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        spreadsheet_key = "1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg"
        return client.open_by_key(spreadsheet_key)
    except:
        return None

def make_match_key(name, opt):
    return str(name).replace(" ", "").upper() + str(opt).replace(" ", "").upper()

def save_reorder_data(new_work_df):
    """기존 타 업체 데이터를 유지하며 현재 데이터만 누적 저장"""
    try:
        spreadsheet = get_sheet()
        sheet = spreadsheet.sheet1
        gs_data = pd.DataFrame(sheet.get_all_records())
        new_work_df['k_tmp'] = new_work_df.apply(lambda r: make_match_key(r['상품명'], r['옵션']), axis=1)
        if gs_data.empty:
            final_df = new_work_df.drop(columns=['k_tmp'])
        else:
            gs_data['k_tmp'] = gs_data.apply(lambda r: make_match_key(r['상품명'], r['옵션']), axis=1)
            old_others = gs_data[~gs_data['k_tmp'].isin(new_work_df['k_tmp'])].copy()
            final_df = pd.concat([old_others, new_work_df], ignore_index=True)
            final_df = final_df.drop(columns=['k_tmp'])
        sheet.clear()
        sheet.update([final_df.columns.values.tolist()] + final_df.fillna(0).values.tolist())
        return True
    except: return False

def save_history_to_gsheet(df, log_type="발주"):
    try:
        spreadsheet = get_sheet()
        try: hist_sheet = spreadsheet.worksheet("history")
        except:
            hist_sheet = spreadsheet.add_worksheet(title="history", rows="1000", cols="20")
            hist_sheet.append_row(["저장시간", "구분", "상품명", "옵션", "수량"])
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        rows_to_add = [[now_str, log_type] + [str(x) for x in row] for row in df.values.tolist()]
        hist_sheet.append_rows(rows_to_add)
        return True
    except: return False

def load_history_from_gsheet():
    try:
        spreadsheet = get_sheet()
        hist_sheet = spreadsheet.worksheet("history")
        return pd.DataFrame(hist_sheet.get_all_records())
    except: return pd.DataFrame()

def find_idx(cols, target_keywords):
    for keyword in target_keywords:
        for i, col in enumerate(cols):
            if keyword in str(col): return i
    return 0

def safe_num(val):
    res = pd.to_numeric(val, errors='coerce')
    return res.fillna(0) if hasattr(res, 'fillna') else (0 if pd.isna(res) else res)

# --- [2. 앱 설정 및 세션 초기화] ---
st.set_page_config(layout="wide", page_title="재고 관리 시스템")
st.title("📦 통합 재고 관리 시스템")

if "extra_order_dict" not in st.session_state: st.session_state.extra_order_dict = {}
if 'analyzed' not in st.session_state: st.session_state.analyzed = False

tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 관리"])

with tab1:
    st.subheader("📁 데이터 업로드 (제작상품)")
    if st.button("🔄 데이터 초기화"):
        for key in ['df_raw', 'analyzed', 'last_filename', 'extra_order_dict']:
            if key in st.session_state: del st.session_state[key]
        st.rerun()

    uploaded_file = st.file_uploader("엑셀/CSV 파일을 선택하세요", type=['xlsx', 'xls', 'csv'], key="prod_upload")

    if uploaded_file is not None:
        if 'df_raw' not in st.session_state or st.session_state.get('last_filename') != uploaded_file.name:
            df_new = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
            df_new.columns = df_new.columns.str.strip()
            try:
                sheet = get_sheet().sheet1
                gs_data = pd.DataFrame(sheet.get_all_records())
                t_item = next((c for c in df_new.columns if '상품명' in c), df_new.columns[0])
                t_opt = next((c for c in df_new.columns if '옵션' in c), df_new.columns[1])
                if not gs_data.empty:
                    df_new['k_tmp'] = df_new.apply(lambda r: make_match_key(r[t_item], r[t_opt]), axis=1)
                    gs_data['k_tmp'] = gs_data.apply(lambda r: make_match_key(r['상품명'], r['옵션']), axis=1)
                    reorder_map = gs_data.set_index('k_tmp')['리오더 수량'].to_dict()
                    df_new['리오더 수량'] = df_new['k_tmp'].map(reorder_map).fillna(0).astype(int)
                    df_new.drop(columns=['k_tmp'], inplace=True)
                else: df_new['리오더 수량'] = 0
            except: df_new['리오더 수량'] = 0
            st.session_state.df_raw = df_new
            st.session_state.last_filename = uploaded_file.name
            st.rerun()

    if st.session_state.get('df_raw') is not None:
        df_curr = st.session_state.df_raw
        cols = df_curr.columns.tolist()
        st.subheader("⚙️ 1단계: 매핑 설정")
        c1, c2 = st.columns(2)
        sold_out = c1.selectbox("품절 여부", cols, index=find_idx(cols, ['품절']))
        vendor = c1.selectbox("공급처", cols, index=find_idx(cols, ['공급처']))
        item = c1.selectbox("상품명", cols, index=find_idx(cols, ['상품명']))
        option = c1.selectbox("옵션", cols, index=find_idx(cols, ['옵션']))
        avail = c2.selectbox("가용재고", cols, index=find_idx(cols, ['가용재고']))
        t3day = c2.selectbox("3일 발주합계", cols, index=find_idx(cols, ['3일']))
        t7day = c2.selectbox("7일 발주합계", cols, index=find_idx(cols, ['7일']))

        st.subheader("🚀 2~3단계: 분석 설정")
        lt = st.number_input("리드타임 (일)", value=10)
        ss = st.number_input("안전재고 (일 수)", value=7)
        if st.button("📊 분석 실행", use_container_width=True):
            st.session_state.analyzed = True
            st.rerun()

    if st.session_state.get('analyzed'):
        st.divider()
        st.subheader("📊 4단계: 데이터 편집 및 재고 관리")
        df_work = st.session_state.df_raw.copy()
        
        # 과거 입고 데이터 로드 및 매핑
        hist_df = load_history_from_gsheet()
        target_date = st.date_input("🗓️ 입고 기록 확인 날짜", datetime.now()).strftime("%Y-%m-%d")
        
        df_work['k_tmp'] = df_work.apply(lambda r: make_match_key(r[item], r[option]), axis=1)
        df_work['과거 리오더입고'] = 0
        if not hist_df.empty and '저장시간' in hist_df.columns:
            hist_df['날짜'] = hist_df['저장시간'].str.split(' ').str[0]
            in_hist = hist_df[(hist_df['날짜'] == target_date) & (hist_df['구분'] == '입고')]
            if not in_hist.empty:
                in_hist['k_tmp'] = in_hist.apply(lambda r: make_match_key(r['상품명'], r['옵션']), axis=1)
                in_map = in_hist.groupby('k_tmp')['수량'].sum().to_dict()
                df_work['과거 리오더입고'] = df_work['k_tmp'].map(in_map).fillna(0).astype(int)

        # 수치 계산
        v7, v3 = safe_num(df_work[t7day]), safe_num(df_work[t3day])
        df_work['일판매량'] = (v7 / 7 if v7.sum() > 0 else v3 / 3).round(0).astype(int)
        df_work['권장발주량'] = ((df_work['일판매량'] * (lt + ss)) - (safe_num(df_work[avail]) + safe_num(df_work['리오더 수량']))).clip(lower=0).astype(int)
        
        df_work['리오더입고수량'] = 0
        display_4 = [vendor, item, option, avail, "리오더 수량", "리오더입고수량", "과거 리오더입고", t3day, "일판매량", "권장발주량"]
        edited_4 = st.data_editor(df_work[display_4], use_container_width=True, hide_index=True, key="ed4")
        
        if st.button("💾 리오더 현황 시트 저장"):
            for idx, row in edited_4.iterrows():
                orig_idx = df_work.index[idx]
                st.session_state.df_raw.at[orig_idx, "리오더 수량"] = int(row["리오더 수량"])
                if int(row["리오더입고수량"]) > 0:
                    st.session_state.df_raw.at[orig_idx, "리오더 수량"] = max(0, int(row["리오더 수량"]) - int(row["리오더입고수량"]))
                    save_history_to_gsheet(pd.DataFrame([[row[item], row[option], row["리오더입고수량"]]], columns=['상품명', '옵션', '수량']), "입고")
            save_reorder_data(st.session_state.df_raw[[item, option, '리오더 수량']].rename(columns={item:'상품명', option:'옵션'}))
            st.rerun()

        st.divider()
        st.subheader("📋 5단계: 최종 발주 리스트 요약")
        df_final = df_work.copy()
        df_final['추가발주수량'] = df_final['k_tmp'].map(st.session_state.extra_order_dict).fillna(0).astype(int)
        df_final['최종발주량'] = df_final['권장발주량'] + df_final['추가발주수량']

        def get_status(r):
            total = safe_num(r[avail]) + safe_num(r['리오더 수량'])
            if r['일판매량'] > 0:
                if total < (r['일판매량'] * 3): return "🚨 긴급"
                if total < (r['일판매량'] * 5): return "⚠️ 주의"
            return "✅ 정상"
        df_final['상태'] = df_final.apply(get_status, axis=1)
        
        display_5 = ["상태", item, option, vendor, avail, "리오더 수량", "추가발주수량", "과거 리오더입고", "권장발주량", "최종발주량"]
        edited_5 = st.data_editor(df_final[display_5], use_container_width=True, hide_index=True, key="ed5")
        
        if st.button("💾 최종 발주 기록 저장"):
            to_save = edited_5[edited_5['최종발주량'].astype(int) > 0]
            if not to_save.empty:
                save_history_to_gsheet(to_save[[item, option, '최종발주량']].rename(columns={item:'상품명', option:'옵션', '최종발주량':'수량'}), "발주")
                st.success("발주 기록 저장 완료!")

# --- [🌙 탭 2: 동대문 사입 관리] ---
with tab2:
    st.subheader("🌙 동대문 사입 관리")
    st.info("여기에 사입 관리 로직을 추가하세요.")
