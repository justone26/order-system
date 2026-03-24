import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta, timezone
import time

# 한국 시간(KST) 설정
KST = timezone(timedelta(hours=9))

# --- [1. 구글 시트 연결 및 데이터 로드 함수] ---
def get_sheet():
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        spreadsheet_key = "1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg"
        return client.open_by_key(spreadsheet_key)
    except Exception as e:
        st.error(f"🚨 구글 시트 연결 실패: {e}"); return None

def load_history_from_gsheet():
    try:
        sh = get_sheet()
        if sh:
            ws = sh.worksheet("history") 
            data = ws.get_all_records()
            return pd.DataFrame(data)
        return pd.DataFrame()
    except: return pd.DataFrame()

def save_reorder_data(new_work_df):
    try:
        spreadsheet = get_sheet()
        if not spreadsheet: return False
        sheet = spreadsheet.sheet1
        raw_gs_data = sheet.get_all_records()
        gs_df = pd.DataFrame(raw_gs_data) if raw_gs_data else pd.DataFrame(columns=['상품명', '옵션', '리오더 수량'])
        def make_key(df_in):
            return df_in['상품명'].astype(str).str.strip().str.replace(" ", "").str.upper() + \
                   df_in['옵션'].astype(str).str.strip().str.replace(" ", "").str.upper()
        gs_df['match_key'] = make_key(gs_df) if not gs_df.empty else ""
        new_work_df['match_key'] = make_key(new_work_df)
        if not gs_df.empty: gs_df['리오더 수량'] = pd.to_numeric(gs_df['리오더 수량'], errors='coerce').fillna(0)
        for _, row in new_work_df.iterrows():
            target_key = row['match_key']
            if target_key in gs_df['match_key'].values:
                gs_df.loc[gs_df['match_key'] == target_key, '리오더 수량'] += row['리오더 수량']
            else:
                new_item = pd.DataFrame([{'상품명': row['상품명'], '옵션': row['옵션'], '리오더 수량': row['리오더 수량']}])
                gs_df = pd.concat([gs_df, new_item], ignore_index=True)
        final_df = gs_df.drop(columns=['match_key'], errors='ignore').fillna(0).drop_duplicates(subset=['상품명', '옵션'], keep='last')
        sheet.clear()
        sheet.update([final_df.columns.values.tolist()] + final_df.values.tolist())
        return True
    except Exception as e: st.error(f"⚠️ 데이터 저장 오류: {e}"); return False

def save_history_to_gsheet(df, log_type="입고"):
    try:
        spreadsheet = get_sheet()
        if not spreadsheet: return False
        try: hist_sheet = spreadsheet.worksheet("history")
        except:
            hist_sheet = spreadsheet.add_worksheet(title="history", rows="1000", cols="20")
            hist_sheet.append_row(["저장시간", "구분", "상품명", "옵션", "수량"])
        now_str = datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S")
        rows_to_add = [[now_str, log_type] + [str(x) for x in row] for row in df.values.tolist()]
        if rows_to_add: hist_sheet.append_rows(rows_to_add); return True
        return False
    except Exception as e: st.error(f"히스토리 저장 실패: {e}"); return False

def find_idx(cols, target_keywords):
    for keyword in target_keywords:
        for i, col in enumerate(cols):
            if keyword in str(col): return i
    return 0

def make_match_key(item, opt):
    return str(item).strip().replace(" ", "").upper() + str(opt).strip().replace(" ", "").upper()

# --- [2. 앱 설정] ---
st.set_page_config(layout="wide", page_title="저스트원 재고관리")
st.title("🏭 저스트원 통합 재고 관리 시스템")

if "add_order_dict" not in st.session_state: st.session_state.add_order_dict = {}
if 'analyzed' not in st.session_state: st.session_state.analyzed = False
if 'df_raw' not in st.session_state: st.session_state.df_raw = None

tab1, tab2 = st.tabs(["✂️ 제작 상품 관리", "🌙 동대문 상품 관리"])

# --- [탭 1: 제작 상품 관리] ---
with tab1:
    uploaded_file = st.file_uploader("엑셀 파일을 올려주세요", type=['xlsx', 'xls', 'csv'], key="t1_up")
    
    if st.button("📂 구글 시트 데이터 로드", use_container_width=True):
        spreadsheet = get_sheet()
        if spreadsheet:
            sheet = spreadsheet.get_worksheet(0)
            raw_data = sheet.get_all_values()
            if len(raw_data) > 1:
                st.session_state.df_raw = pd.DataFrame(raw_data[1:], columns=[str(h).strip() for h in raw_data[0]])
                st.session_state.analyzed = True
                st.success("✅ 로드 완료!"); time.sleep(0.5); st.rerun()

    if uploaded_file and (st.session_state.df_raw is None or st.session_state.get('last_fn') != uploaded_file.name):
        df_new = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
        df_new.columns = df_new.columns.str.strip()
        st.session_state.df_raw = df_new
        st.session_state.last_fn = uploaded_file.name
        st.rerun()

    if st.session_state.df_raw is not None:
        df_curr = st.session_state.df_raw
        cols = df_curr.columns.tolist()
        
        st.subheader("⚙️ 3단계: 매핑 설정")
        c_l, c_r = st.columns(2)
        with c_l:
            sel_sold_out = st.selectbox("품절 여부", cols, index=find_idx(cols, ['품절']))
            sel_vendor = st.selectbox("공급처", cols, index=find_idx(cols, ['공급처']))
            sel_v_item = st.selectbox("공급처 상품명", cols, index=find_idx(cols, ['공급처상품명']))
            sel_item = st.selectbox("상품명", cols, index=find_idx(cols, ['상품명']))
            sel_option = st.selectbox("옵션", cols, index=find_idx(cols, ['옵션']))
        with c_r:
            sel_stock = st.selectbox("정상재고", cols, index=find_idx(cols, ['정상재고']))
            sel_avail = st.selectbox("가용재고", cols, index=find_idx(cols, ['가용재고']))
            sel_t3day = st.selectbox("3일 발주합계", cols, index=find_idx(cols, ['3일']))
            sel_t7day = st.selectbox("7일 발주합계", cols, index=find_idx(cols, ['7일', '1주']))
            lt = st.number_input("리드타임 (일)", value=7)
            ss = st.number_input("안전재고 (일 수)", value=3)

        # --- [4단계: 데이터 편집 및 리오더 차감] ---
        st.divider()
        st.subheader("📊 4단계: 데이터 편집 및 재고 관리")
        df_work = df_curr.copy()
        for c in [sel_stock, sel_avail, sel_t3day, sel_t7day]:
            df_work[c] = pd.to_numeric(df_work[c], errors='coerce').fillna(0).astype(int)
        
        if "리오더 수량" not in df_work.columns: df_work["리오더 수량"] = 0
        df_work["리오더 수량"] = pd.to_numeric(df_work["리오더 수량"], errors='coerce').fillna(0).astype(int)

        df_work['일판매량'] = (df_work[sel_t7day] / 7 if df_work[sel_t7day].sum() > 0 else df_work[sel_t3day] / 3).round(0).astype(int)
        df_work['권장발주량'] = ((df_work['일판매량'] * (lt + ss)) - (df_work[sel_avail] + df_work['리오더 수량'])).clip(lower=0).astype(int)

        f_c1, f_c2, f_c3 = st.columns([2, 1, 1])
        search_q = f_c1.text_input("🔍 상품명 검색", key="search_4")
        filter_m = f_c2.selectbox("품절 필터", ["전체보기", "정상만", "품절만"], index=1)
        
        if filter_m == "정상만": df_work = df_work[~df_work[sel_sold_out].astype(str).str.contains('품절', na=False)]
        elif filter_m == "품절만": df_work = df_work[df_work[sel_sold_out].astype(str).str.contains('품절', na=False)]
        if search_q: df_work = df_work[df_work[sel_item].astype(str).str.contains(search_q, case=False)]

        df_disp = df_work.rename(columns={sel_sold_out: "품절", sel_vendor: "공급쳐", sel_item: "상품명", sel_option: "옵션", sel_stock: "정상재고", sel_avail: "가용재고"})
        
        with st.form("form_4"):
            edited_4 = st.data_editor(df_disp, use_container_width=True, hide_index=True, key="editor_4")
            if st.form_submit_button("💾 입고량 반영 및 저장", use_container_width=True, type="primary"):
                edits = st.session_state["editor_4"].get("edited_rows", {})
                if edits:
                    for r_idx_str, change in edits.items():
                        orig_idx = df_work.index[int(r_idx_str)]
                        if "리오더 입고수량" in change: # 사장님 차감 로직
                            in_qty = int(change["리오더 입고수량"])
                            st.session_state.df_raw.at[orig_idx, "리오더 수량"] = max(0, int(st.session_state.df_raw.at[orig_idx, "리오더 수량"]) - in_qty)
                    save_reorder_data(st.session_state.df_raw[[sel_item, sel_option, '리오더 수량']].rename(columns={sel_item:'상품명', sel_option:'옵션'}))
                    st.success("✅ 저장 완료!"); time.sleep(0.5); st.rerun()

        # --- [5단계: 최종 발주 요약] ---
        st.divider()
        st.subheader("📋 5단계: 최종 발주 리스트 요약")
        df_5 = st.session_state.df_raw.copy()
        # (중간 계산 로직 사장님 원본 그대로...)
        def get_status(r):
            s_sum = pd.to_numeric(r[sel_avail], errors='coerce') + pd.to_numeric(r['리오더 수량'], errors='coerce')
            d = (pd.to_numeric(r[sel_t7day], errors='coerce') / 7).round(0)
            if d > 0:
                if s_sum < (d * 3): return "🚨 긴급"
                if s_sum < (d * 5): return "⚠️ 주의"
            return "✅ 정상"
        df_5['상태'] = df_5.apply(get_status, axis=1)
        st.dataframe(df_5[df_5['상태'] != "✅ 정상"].sort_values('상태'), use_container_width=True)

        # --- [6단계: 히스토리 내역] ---
        st.divider()
        st.subheader("📜 6단계: 전체 히스토리 내역")
        df_hist = load_history_from_gsheet()
        if not df_hist.empty:
            h_date = st.date_input("🗓️ 조회 날짜", datetime.now(KST).date())
            st.dataframe(df_hist, use_container_width=True)

# --- [탭 2: 동대문 상품 관리 (사장님 가중치 로직)] ---
with tab2:
    st.subheader("🌙 동대문 사입 및 미납 관리")
    dong_file = st.file_uploader("동대문 엑셀 업로드", key="dong_up")
    if dong_file:
        df_dong = pd.read_excel(dong_file)
        df_dong['판매수량'] = (df_dong['정상재고'] - df_dong['가용재고']).clip(lower=0)
        df_dong['가중율'] = df_dong['판매수량'].apply(lambda n: 2.0 if n >= 10 else (1.5 if n >= 6 else 1.0))
        df_dong['발주수량'] = (df_dong['판매수량'] * df_dong['가중율']).astype(int)
        st.data_editor(df_dong, use_container_width=True)
