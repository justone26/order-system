import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- [1. 공통 함수 및 구글 시트 연동] ---
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
    except Exception as e:
        st.error(f"구글 시트 저장 실패: {e}")

def save_history_to_gsheet(df):
    try:
        spreadsheet = get_sheet()
        try: hist_sheet = spreadsheet.worksheet("history")
        except: 
            hist_sheet = spreadsheet.add_worksheet(title="history", rows="1000", cols="20")
            hist_sheet.append_row(["저장시간"] + df.columns.tolist())
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        rows_to_add = [[now_str] + [str(x) for x in row] for row in df.values.tolist()]
        hist_sheet.append_rows(rows_to_add)
        return True
    except: return False

def load_history_from_gsheet():
    try:
        spreadsheet = get_sheet()
        hist_sheet = spreadsheet.worksheet("history")
        data = hist_sheet.get_all_records()
        return pd.DataFrame(data)
    except: return pd.DataFrame()

# [추가] 특정 상품의 과거 입고 내역 요약 가져오기
def get_in_qty_logs():
    try:
        df_h = load_history_from_gsheet()
        if df_h.empty: return {}
        # 리오더입고수량이 있는 데이터만 필터
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
    st.divider()

    if uploaded_file is not None:
        if 'df_raw' not in st.session_state or st.session_state.get('last_filename') != uploaded_file.name:
            df_new = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
            df_new.columns = df_new.columns.str.strip()
            df_new = df_new.loc[:, ~df_new.columns.duplicated()] 

            # [해결 1] 초기화 방지: 업로드 즉시 구글 시트 리오더 수량 병합
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
            st.session_state.in_logs = get_in_qty_logs() # 과거 입고 기록 로드
            st.rerun()

    if st.session_state.get('df_raw') is not None:
        df_current = st.session_state.df_raw
        cols = df_current.columns.tolist()

        # 1단계 매핑
        st.subheader("⚙️ 1단계: 매핑 설정")
        c1, c2 = st.columns(2)
        sold_out = c1.selectbox("품절 여부", cols, index=find_idx(cols, ['품절']))
        vendor = c1.selectbox("공급처", cols, index=find_idx(cols, ['공급처']))
        item = c1.selectbox("상품명", cols, index=find_idx(cols, ['상품명']))
        option = c1.selectbox("옵션", cols, index=find_idx(cols, ['옵션']))
        vendor_item = c1.selectbox("공급처 상품명", cols, index=find_idx(cols, ['공급처상품명']))
        reg_date = c2.selectbox("등록일", cols, index=find_idx(cols, ['등록일']))
        stock = c2.selectbox("정상재고", cols, index=find_idx(cols, ['정상재고']))
        avail = c2.selectbox("가용재고", cols, index=find_idx(cols, ['가용재고']))
        t3day = c2.selectbox("3일 발주합계", cols, index=find_idx(cols, ['3일']))
        t1week = c2.selectbox("7일 발주합계", cols, index=find_idx(cols, ['7일', '1주']))

        # 2-3단계 설정
        st.subheader("⚙️ 2~3단계: 분석 설정")
        col_lt, col_ss = st.columns(2)
        lead_time = col_lt.number_input("리드타임 (일)", value=10)
        safety_stock = col_ss.number_input("안전재고 (일 수)", value=7)
        
        if st.button("🚀 분석 실행"):
            st.session_state.analyzed = True
            st.rerun()

        if st.session_state.analyzed:
            st.subheader("📊 4단계: 데이터 편집 및 재고 관리")
            f1, f2 = st.columns([3, 1])
            search_query = f1.text_input("🔍 상품명 검색", key="prod_search_input")
            filter_mode = f2.selectbox("품절 필터", ["전체보기", "정상만", "품절만"], index=1)
            
            df_working = st.session_state.df_raw.copy()
            
            # 계산 로직
            v_avail = pd.to_numeric(df_working[avail], errors='coerce').fillna(0)
            v_reorder = pd.to_numeric(df_working['리오더 수량'], errors='coerce').fillna(0)
            v_3day = pd.to_numeric(df_working[t3day], errors='coerce').fillna(0)
            df_working["일판매량"] = (v_3day / 3).round(0).astype(int)
            
            # [해결 2] 품절건 제외 로직 적용
            df_working["권장발주량"] = ((df_working["일판매량"] * (lead_time + safety_stock)) - (v_avail + v_reorder)).clip(lower=0).round(0).astype(int)
            df_working.loc[df_working[sold_out].astype(str).str.contains('품절', na=False), "권장발주량"] = 0

            # 필터링
            if filter_mode == "정상만": df_working = df_working[~df_working[sold_out].astype(str).str.contains('품절', na=False)]
            elif filter_mode == "품절만": df_working = df_working[df_working[sold_out].astype(str).str.contains('품절', na=False)]
            if search_query: df_working = df_working[df_working[item].astype(str).str.contains(search_query, case=False, na=False)]

            # [해결 3] 최근 입고 내역 매칭 셀 추가
            def get_log_str(row):
                key = (str(row[item]).strip(), str(row[option]).strip())
                return st.session_state.in_logs.get(key, "-")
            df_working['최근입고기록'] = df_working.apply(get_log_str, axis=1)

            if "리오더입고수량" not in df_working.columns: df_working["리오더입고수량"] = 0

            def auto_save_and_update():
                if "main_editor" in st.session_state and st.session_state["main_editor"]["edited_rows"]:
                    changes = st.session_state["main_editor"]["edited_rows"]
                    for row_idx_str, change in changes.items():
                        row_idx = int(row_idx_str)
                        orig_idx = df_working.index[row_idx]
                        if "리오더 수량" in change:
                            st.session_state.df_raw.at[orig_idx, "리오더 수량"] = int(change["리오더 수량"])
                        if "리오더입고수량" in change:
                            in_qty = int(change["리오더입고수량"])
                            curr = st.session_state.df_raw.at[orig_idx, "리오더 수량"]
                            st.session_state.df_raw.at[orig_idx, "리오더 수량"] = max(0, curr - in_qty)
                            # 입고 시 history에 기록
                            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            sh = get_sheet()
                            sh.worksheet("history").append_row([now, str(df_working.at[orig_idx, item]), str(df_working.at[orig_idx, option]), in_qty])
                    
                    # 시트 Sheet1 저장
                    save_df = st.session_state.df_raw[[item, option, '리오더 수량']].copy()
                    save_df.columns = ['상품명', '옵션', '리오더 수량']
                    save_reorder_data(save_df)
                    st.session_state.in_logs = get_in_qty_logs()
                    st.rerun()

            display_cols_4 = [sold_out, item, option, vendor_item, stock, avail, "리오더 수량", "리오더입고수량", "최근입고기록", "일판매량", t3day, "권장발주량"]
            final_target_4 = [c for c in display_cols_4 if c in df_working.columns or c in ["최근입고기록", "리오더입고수량"]]

            st.data_editor(
                df_working[final_target_4], use_container_width=True, height=400, key="main_editor",
                on_change=auto_save_and_update,
                column_config={
                    "리오더 수량": st.column_config.Column("📝 리오더 수량"),
                    "리오더입고수량": st.column_config.Column("➕ 입고입력"),
                    "최근입고기록": st.column_config.Column("📅 최근입고일(수량)", width="medium")
                }
            )

            # --- 5단계: 요약 ---
            st.subheader("📋 5단계: 최종 발주 리스트 요약")
            to_order = st.session_state.df_raw.copy()
            # ... (기존 5단계 로직 동일 유지) ...
            v_3day_val = pd.to_numeric(to_order[t3day], errors='coerce').fillna(0)
            to_order['일판매량'] = (v_3day_val / 3).round(0).astype(int)
            to_order['권장발주량'] = ((to_order['일판매량'] * (lead_time + safety_stock)) - (pd.to_numeric(to_order[avail], errors='coerce').fillna(0) + to_order['리오더 수량'])).clip(lower=0).round(0).astype(int)
            to_order.loc[to_order[sold_out].astype(str).str.contains('품절', na=False), '권장발주량'] = 0
            
            to_order = to_order[to_order['권장발주량'] > 0]
            if not to_order.empty:
                st.dataframe(to_order[[item, option, vendor_item, avail, "리오더 수량", "권장발주량"]], use_container_width=True)
                csv = to_order.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
                st.download_button("📥 엑셀 다운로드", csv, f"발주_{datetime.now().strftime('%m%d_%H%M')}.csv", "text/csv")
            else: st.info("💡 발주할 상품이 없습니다.")

            # 6단계 확인
            st.divider()
            st.subheader("📜 6단계: 과거 데이터 확인")
            if st.button("🔄 기록 불러오기"): st.session_state.db_history = load_history_from_gsheet()
            if 'db_history' in st.session_state and not st.session_state.db_history.empty:
                st.dataframe(st.session_state.db_history.sort_values(by='저장시간', ascending=False), use_container_width=True)

# --- [🌙 탭 2: 동대문 사입 관리] ---
with tab2:
    st.subheader("🌙 동대문 사입 및 미납 관리")
    # 기존 코드 그대로 유지
    dong_file = st.file_uploader("동대문 주문 리스트 업로드", type=['xlsx', 'csv'], key="dong_tab_upload")
    if dong_file:
        if "last_file_name" not in st.session_state or st.session_state.last_file_name != dong_file.name:
            df = pd.read_excel(dong_file) if not dong_file.name.endswith('.csv') else pd.read_csv(dong_file)
            st.session_state.df_dong_current = df
            st.session_state.last_file_name = dong_file.name
        st.data_editor(st.session_state.df_dong_current, use_container_width=True, key="dong_editor")
