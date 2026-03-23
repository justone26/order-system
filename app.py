import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- [1. 공통 함수 정의] ---
def get_sheet():
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        spreadsheet_key = "1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg"
        return client.open_by_key(spreadsheet_key)
    except: return None

def save_reorder_data(df):
    try:
        sheet = get_sheet().sheet1
        sheet.clear()
        sheet.update([df.columns.values.tolist()] + df.values.tolist())
    except: pass

def save_history_to_gsheet(df):
    try:
        spreadsheet = get_sheet()
        try: hist_sheet = spreadsheet.worksheet("history")
        except: 
            hist_sheet = spreadsheet.add_worksheet(title="history", rows="1000", cols="20")
            hist_sheet.append_row(["저장시간", "상품명", "옵션", "리오더입고수량"])
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        rows_to_add = [[now_str] + [str(x) for x in row] for row in df.values.tolist()]
        hist_sheet.append_rows(rows_to_add)
        return True
    except: return False

def load_history_from_gsheet():
    try:
        spreadsheet = get_sheet()
        hist_sheet = spreadsheet.worksheet("history")
        return pd.DataFrame(hist_sheet.get_all_records())
    except: return pd.DataFrame()

def get_in_qty_logs():
    try:
        df_h = load_history_from_gsheet()
        if df_h.empty: return {}
        if '리오더입고수량' in df_h.columns:
            df_in = df_h[pd.to_numeric(df_h['리오더입고수량'], errors='coerce') > 0].copy()
            log_dict = {}
            for (name, opt), group in df_in.groupby(['상품명', '옵션']):
                recent = group.sort_values(by='저장시간', ascending=False).head(3)
                logs = [f"{str(r['저장시간'])[5:10]}({int(r['리오더입고수량'])}개)" for _, r in recent.iterrows()]
                log_dict[(str(name).strip(), str(opt).strip())] = " / ".join(logs)
            return log_dict
    except: return {}

def find_idx(cols, target_keywords):
    for keyword in target_keywords:
        for i, col in enumerate(cols):
            if keyword in str(col): return i
    return 0

# --- [2. 앱 설정 및 탭 구성] ---
st.set_page_config(layout="wide", page_title="재고 관리 시스템")
st.title("📦 통합 재고 관리 시스템")

tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 관리"])

# --- [🏭 탭 1: 제작 상품 관리] ---
with tab1:
    if 'analyzed' not in st.session_state: st.session_state.analyzed = False
    if 'in_logs' not in st.session_state: st.session_state.in_logs = {}

    st.subheader("📁 데이터 업로드 (제작상품)")
    if st.button("🔄 제작상품 데이터 초기화"):
        st.session_state.clear()
        st.rerun()

    uploaded_file = st.file_uploader("엑셀/CSV 파일을 선택하세요", type=['xlsx', 'xls', 'csv'], key="prod_upload")

    if uploaded_file is not None:
        if 'df_raw' not in st.session_state or st.session_state.get('last_filename') != uploaded_file.name:
            df_new = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
            df_new.columns = df_new.columns.str.strip()
            df_new = df_new.loc[:, ~df_new.columns.duplicated()] 

            try:
                sheet = get_sheet().sheet1
                gs_data = pd.DataFrame(sheet.get_all_records())
                if not gs_data.empty and '리오더 수량' in gs_data.columns:
                    t_item = next((c for c in df_new.columns if '상품명' in c), df_new.columns[0])
                    t_opt = next((c for c in df_new.columns if '옵션' in c), df_new.columns[1])
                    df_new['match_key'] = df_new[t_item].astype(str).str.strip() + df_new[t_opt].astype(str).str.strip()
                    gs_data['match_key'] = gs_data['상품명'].astype(str).str.strip() + gs_data['옵션'].astype(str).str.strip()
                    reorder_map = gs_data.set_index('match_key')['리오더 수량'].to_dict()
                    df_new['리오더 수량'] = df_new['match_key'].map(reorder_map).fillna(0).astype(int)
                    df_new.drop(columns=['match_key'], inplace=True)
                else: df_new['리오더 수량'] = 0
            except: df_new['리오더 수량'] = 0

            st.session_state.df_raw = df_new
            st.session_state.last_filename = uploaded_file.name
            st.session_state.in_logs = get_in_qty_logs()
            st.rerun()

    if st.session_state.get('df_raw') is not None:
        df_curr = st.session_state.df_raw
        cols = df_curr.columns.tolist()

        st.subheader("⚙️ 1단계: 매핑 설정")
        c1, c2 = st.columns(2)
        sold_out = c1.selectbox("품절 여부", cols, index=find_idx(cols, ['품절']))
        item = c1.selectbox("상품명", cols, index=find_idx(cols, ['상품명']))
        option = c1.selectbox("옵션", cols, index=find_idx(cols, ['옵션']))
        vendor_item = c1.selectbox("공급처 상품명", cols, index=find_idx(cols, ['공급처상품명']))
        reg_date = c2.selectbox("등록일", cols, index=find_idx(cols, ['등록일']))
        stock = c2.selectbox("정상재고", cols, index=find_idx(cols, ['정상재고']))
        avail = c2.selectbox("가용재고", cols, index=find_idx(cols, ['가용재고']))
        t3day = c2.selectbox("3일 발주합계", cols, index=find_idx(cols, ['3일']))

        st.subheader("🚀 2~3단계: 분석 설정")
        l1, l2 = st.columns(2)
        lt = l1.number_input("리드타임 (일)", value=10)
        ss = l2.number_input("안전재고 (일 수)", value=7)
        if st.button("📊 분석 실행", use_container_width=True):
            st.session_state.analyzed = True
            st.rerun()

        if st.session_state.analyzed:
            st.subheader("📊 4단계: 데이터 편집 및 재고 관리")
            df_work = st.session_state.df_raw.copy()
            
            # 필터 레이아웃 복구 (검색창 + 품절 필터)
            f_col1, f_col2 = st.columns([3, 1])
            s_query = f_col1.text_input("🔍 상품명 검색")
            # [핵심 수정] 품절 필터 복구 및 '정상만' 기본값 설정
            f_mode = f_col2.selectbox("품절 필터", ["전체보기", "정상만", "품절만"], index=1)

            # 계산
            v_av = pd.to_numeric(df_work[avail], errors='coerce').fillna(0)
            v_re = pd.to_numeric(df_work['리오더 수량'], errors='coerce').fillna(0)
            df_work['일판매량'] = (pd.to_numeric(df_work[t3day], errors='coerce').fillna(0) / 3).round(0).astype(int)
            df_work['권장발주량'] = ((df_work['일판매량'] * (lt + ss)) - (v_av + v_re)).clip(lower=0).astype(int)
            df_work.loc[df_work[sold_out].astype(str).str.contains('품절', na=False), '권장발주량'] = 0

            # 필터링 적용
            if f_mode == "정상만": df_work = df_work[~df_work[sold_out].astype(str).str.contains('품절', na=False)]
            elif f_mode == "품절만": df_work = df_work[df_work[sold_out].astype(str).str.contains('품절', na=False)]
            if s_query: df_work = df_work[df_work[item].astype(str).str.contains(s_query, case=False)]
            
            def match_log(row):
                key = (str(row[item]).strip(), str(row[option]).strip())
                return st.session_state.in_logs.get(key, "-")
            df_work['최근입고기록'] = df_work.apply(match_log, axis=1)
            if "리오더입고수량" not in df_work.columns: df_work["리오더입고수량"] = 0

            def on_main_edit():
                changes = st.session_state["main_editor"]["edited_rows"]
                for r_idx_str, change in changes.items():
                    idx = int(r_idx_str)
                    orig_idx = df_work.index[idx]
                    if "리오더 수량" in change:
                        st.session_state.df_raw.at[orig_idx, "리오더 수량"] = int(change["리오더 수량"])
                    if "리오더입고수량" in change:
                        in_qty = int(change["리오더입고수량"])
                        curr = st.session_state.df_raw.at[orig_idx, "리오더 수량"]
                        st.session_state.df_raw.at[orig_idx, "리오더 수량"] = max(0, curr - in_qty)
                        sh = get_sheet()
                        sh.worksheet("history").append_row([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), str(df_work.at[orig_idx, item]), str(df_work.at[orig_idx, option]), in_qty])
                save_df = st.session_state.df_raw[[item, option, '리오더 수량']].copy()
                save_df.columns = ['상품명', '옵션', '리오더 수량']
                save_reorder_data(save_df)
                st.session_state.in_logs = get_in_qty_logs()
                st.rerun()

            disp_cols = [sold_out, item, option, vendor_item, stock, avail, "리오더 수량", "리오더입고수량", "최근입고기록", "일판매량", "권장발주량"]
            st.data_editor(df_work[disp_cols], use_container_width=True, key="main_editor", on_change=on_main_edit)

            # 5단계 최종 요약
            st.divider()
            st.subheader("📋 5단계: 최종 발주 리스트 요약")
            to_order = df_work.copy()
            def check_status(row):
                v_sl = row['일판매량']
                v_tot = pd.to_numeric(row[avail], errors='coerce') + row['리오더 수량']
                if v_sl > 0 and v_tot < (v_sl * 3): return "🚨 긴급"
                elif v_sl > 0 and v_tot < (v_sl * 5): return "⚠️ 주의"
                return "✅ 정상"
            to_order['상태'] = to_order.apply(check_status, axis=1)
            s_col1, s_col2 = st.columns([1, 1])
            s_filter = s_col1.selectbox("🎯 상태 필터", ["전체보기", "🚨 긴급만 보기", "⚠️ 주의이상 보기"])
            order_mask = (to_order['권장발주량'] > 0) | (to_order['상태'] != "✅ 정상")
            df_final = to_order[order_mask].copy()
            if "🚨" in s_filter: df_final = df_final[df_final['상태'] == "🚨 긴급"]
            elif "⚠️" in s_filter: df_final = df_final[df_final['상태'].str.contains("🚨|⚠️")]

            if not df_final.empty:
                final_cols = ["상태", item, option, vendor_item, avail, "리오더 수량", "권장발주량"]
                edited_final = st.data_editor(df_final[final_cols], use_container_width=True, key="final_order_editor")
                col_btn1, col_btn2 = st.columns(2)
                if col_btn1.button("💾 구글 시트에 최종 기록 저장", use_container_width=True):
                    if save_history_to_gsheet(edited_final): st.success("✅ 히스토리에 저장되었습니다!")
                csv = edited_final.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
                col_btn2.download_button("📥 엑셀 다운로드", csv, f"발주서_{datetime.now().strftime('%m%d')}.csv", use_container_width=True)
            else: st.info("💡 현재 발주가 필요한 상품이 없습니다.")

# --- [🌙 탭 2: 동대문 관리 및 히스토리] ---
with tab2:
    st.subheader("📜 6단계: 입고 히스토리 확인")
    if st.button("🔄 히스토리 불러오기"): st.session_state.db_hist = load_history_from_gsheet()
    if 'db_hist' in st.session_state and not st.session_state.db_hist.empty:
        st.dataframe(st.session_state.db_hist.sort_values(by='저장시간', ascending=False), use_container_width=True)
