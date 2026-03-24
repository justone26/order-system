import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- [1. 기본 함수 설정] ---
def get_sheet():
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        spreadsheet_key = "1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg"
        return client.open_by_key(spreadsheet_key)
    except: return None

def load_history_from_gsheet():
    try:
        spreadsheet = get_sheet()
        hist_sheet = spreadsheet.worksheet("history")
        data = hist_sheet.get_all_records()
        return pd.DataFrame(data)
    except: return pd.DataFrame()

def make_match_key(name, opt):
    return str(name).strip().replace(" ", "").upper() + str(opt).strip().replace(" ", "").upper()

def save_reorder_data(new_work_df):
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

def save_history_to_gsheet(df, log_type="입고"):
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

def find_idx(cols, target_keywords):
    for keyword in target_keywords:
        for i, col in enumerate(cols):
            if keyword in str(col): return i
    return 0

def safe_num(val):
    res = pd.to_numeric(val, errors='coerce')
    if isinstance(res, pd.Series): return res.fillna(0)
    return 0 if pd.isna(res) else res

# --- [2. 앱 초기 설정] ---
st.set_page_config(layout="wide", page_title="저스트원 재고관리")
st.title("🏭 저스트원 통합 재고 관리 시스템")

if "extra_order_dict" not in st.session_state: st.session_state.extra_order_dict = {}
if 'analyzed' not in st.session_state: st.session_state.analyzed = False

tab1, tab2 = st.tabs(["✂️ 제작 상품 관리", "🌙 동대문 상품 관리"])

# --- [탭 1: 제작 상품 관리] ---
with tab1:
    uploaded_file = st.file_uploader("엑셀 파일을 올려주세요", type=['xlsx', 'xls', 'csv'], key="t1_up")
    
    if st.button("📂 구글 시트 데이터 로드", use_container_width=True):
        try:
            sheet = get_sheet().sheet1
            gs_df = pd.DataFrame(sheet.get_all_records())
            if not gs_df.empty:
                st.session_state.df_raw = gs_df.copy()
                st.session_state.analyzed = False
                st.rerun()
        except: st.error("시트 연결 실패")

    if uploaded_file:
        if 'df_raw' not in st.session_state or st.session_state.get('last_fn') != uploaded_file.name:
            df_new = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
            df_new.columns = df_new.columns.str.strip()
            # 리오더 수량 매핑
            try:
                sheet = get_sheet().sheet1
                gs_df = pd.DataFrame(sheet.get_all_records())
                if not gs_df.empty and '리오더 수량' in gs_df.columns:
                    t_item = next((c for c in df_new.columns if '상품명' in c), df_new.columns[0])
                    t_opt = next((c for c in df_new.columns if '옵션' in c), df_new.columns[1])
                    df_new['k_tmp'] = df_new.apply(lambda r: make_match_key(r[t_item], r[t_opt]), axis=1)
                    gs_df['k_tmp'] = gs_df.apply(lambda r: make_match_key(r['상품명'], r['옵션']), axis=1)
                    rmap = gs_df.set_index('k_tmp')['리오더 수량'].to_dict()
                    df_new['리오더 수량'] = df_new['k_tmp'].map(rmap).fillna(0).astype(int)
                    df_new.drop(columns=['k_tmp'], inplace=True)
                else: df_new['리오더 수량'] = 0
            except: df_new['리오더 수량'] = 0
            st.session_state.df_raw = df_new
            st.session_state.last_fn = uploaded_file.name
            st.session_state.analyzed = False
            st.rerun()

    if st.session_state.get('df_raw') is not None:
        df_curr = st.session_state.df_raw
        cols = df_curr.columns.tolist()
        
        st.subheader("⚙️ 3단계: 매핑 설정")
        c_l, c_r = st.columns(2)
        with c_l:
            sold_out = st.selectbox("품절 여부", cols, index=find_idx(cols, ['품절']))
            vendor = st.selectbox("공급처", cols, index=find_idx(cols, ['공급처']))
            v_item = st.selectbox("공급처 상품명", cols, index=find_idx(cols, ['공급처상품명']))
            item = st.selectbox("상품명", cols, index=find_idx(cols, ['상품명']))
            option = st.selectbox("옵션", cols, index=find_idx(cols, ['옵션']))
        with c_r:
            reg_date = st.selectbox("등록일", cols, index=find_idx(cols, ['등록일']))
            stock = st.selectbox("정상재고", cols, index=find_idx(cols, ['정상재고']))
            avail = st.selectbox("가용재고", cols, index=find_idx(cols, ['가용재고']))
            t3day = st.selectbox("3일 발주합계", cols, index=find_idx(cols, ['3일']))
            t7day = st.selectbox("7일 발주합계", cols, index=find_idx(cols, ['7일', '1주']))

        lt = st.number_input("리드타임 (일)", value=7)
        ss = st.number_input("안전재고 (일 수)", value=3)
        if st.button("📊 분석 실행", use_container_width=True):
            st.session_state.analyzed = True
            st.rerun()
            
    if st.session_state.get('analyzed'):
        # --- [4단계: 데이터 편집 및 재고 관리] ---
        st.divider()
        st.subheader("📊 4단계: 데이터 편집 및 재고 관리")
        df_work = st.session_state.df_raw.copy()

        f_c1, f_c2, f_c3 = st.columns([2, 1, 1])
        search_q = f_c1.text_input("🔍 상품명 검색", key="search_v4")
        filter_m = f_c2.selectbox("품절 필터", ["전체보기", "정상만", "품절만"], index=1, key="filter_v4")
        hist_date_4 = f_c3.date_input("🗓️ 입고 기록 확인 날짜", datetime.now(), key="date_v4")

        def simple_key(n): return str(n).strip().replace(" ", "").upper() if not pd.isna(n) else ""
        df_work['unique_key'] = df_work[item].apply(simple_key) + df_work[option].apply(simple_key)

        past_hist = load_history_from_gsheet()
        df_work['과거 리오더입고'] = 0
        df_work['리오더입고수량'] = 0
        
        if not past_hist.empty:
            if '저장시간' in past_hist.columns:
                past_hist['날짜'] = past_hist['저장시간'].astype(str).str.split(' ').str[0]
                target_date_str = hist_date_4.strftime("%Y-%m-%d")
                if '구분' in past_hist.columns:
                    t_hist = past_hist[(past_hist['날짜'] == target_date_str) & (past_hist['구분'] == "입고")].copy()
                    if not t_hist.empty:
                        t_hist['k_tmp'] = t_hist['상품명'].apply(simple_key) + t_hist['옵션'].apply(simple_key)
                        in_map = t_hist.groupby('k_tmp')['수량'].sum().to_dict()
                        df_work['과거 리오더입고'] = df_work['unique_key'].map(in_map).fillna(0).astype(int)

        v7 = safe_num(df_work[t7day]); v3 = safe_num(df_work[t3day])
        df_work['일판매량'] = (v7 / 7 if v7.sum() > 0 else v3 / 3).round(0).astype(int)
        df_work['권장발주량'] = ((df_work['일판매량'] * (lt + ss)) - (safe_num(df_work[avail]) + safe_num(df_work['리오더 수량']))).clip(lower=0).astype(int)

        if filter_m == "정상만": df_work = df_work[~df_work[sold_out].astype(str).str.contains('품절', na=False)]
        elif filter_m == "품절만": df_work = df_work[df_work[sold_out].astype(str).str.contains('품절', na=False)]
        if search_q: df_work = df_work[df_work[item].astype(str).str.contains(search_q, case=False, na=False)]

        display_cols = [sold_out, vendor, v_item, item, option, stock, avail, "리오더 수량", "리오더입고수량", "과거 리오더입고", t3day, "일판매량", "권장발주량"]
        valid_cols = [c for c in display_cols if c in df_work.columns or c in ["리오더입고수량", "과거 리오더입고"]]
        
        def on_edit_4():
            changes = st.session_state["editor_v4"]["edited_rows"]
            for r_idx_str, change in changes.items():
                orig_idx = df_work.index[int(r_idx_str)]
                if "리오더 수량" in change:
                    st.session_state.df_raw.at[orig_idx, "리오더 수량"] = int(change["리오더 수량"])
                if "리오더입고수량" in change:
                    in_qty = int(change["리오더입고수량"])
                    st.session_state.df_raw.at[orig_idx, "리오더 수량"] = max(0, int(st.session_state.df_raw.at[orig_idx, "리오더 수량"]) - in_qty)
                    save_history_to_gsheet(pd.DataFrame([[df_work.at[orig_idx, item], df_work.at[orig_idx, option], in_qty]], columns=['상품명', '옵션', '수량']), log_type="입고")
            save_reorder_data(st.session_state.df_raw[[item, option, '리오더 수량']].rename(columns={item:'상품명', option:'옵션'}))
            st.rerun()

        st.data_editor(df_work[valid_cols], use_container_width=True, key="editor_v4", on_change=on_edit_4, hide_index=True)

        # --- [5단계: 최종 발주 리스트 요약] ---
        st.divider()
        st.subheader("📋 5단계: 최종 발주 리스트 요약")
        c5_1, c5_2 = st.columns([2, 1])
        s_filter = c5_1.selectbox("🎯 상태 필터", ["🚨긴급 + ⚠️주의 우선", "🚨 긴급만 보기", "✅ 정상 포함 전체보기"], index=0, key="s_filter_v5")
        hist_date_5 = c5_2.date_input("🗓️ 입고 기록 확인 날짜 (연동)", value=hist_date_4, key="date_v5")

        to_order = df_work.copy()
        to_order['추가발주수량'] = to_order['unique_key'].map(st.session_state.extra_order_dict).fillna(0).astype(int)
        to_order['최종발주량'] = to_order['권장발주량'].astype(int) + to_order['추가발주수량'].astype(int)

        def get_final_status(r):
            total = safe_num(r[avail]) + safe_num(r['리오더 수량'])
            daily = r['일판매량']
            if daily > 0:
                if total < (daily * 3): return "🚨 긴급"
                if total < (daily * 5): return "⚠️ 주의"
            return "✅ 정상"
        
        to_order['상태'] = to_order.apply(get_final_status, axis=1)
        to_order = to_order.sort_values(by='상태') # 긴급 우선

        if s_filter == "🚨긴급 + ⚠️주의 우선": to_order = to_order[to_order['상태'].isin(["🚨 긴급", "⚠️ 주의"]) | (to_order['최종발주량'] > 0)]
        elif s_filter == "🚨 긴급만 보기": to_order = to_order[to_order['상태'] == "🚨 긴급"]

        disp_final = ["상태", item, option, vendor, avail, "리오더 수량", "추가발주수량", "과거 리오더입고", "권장발주량", "최종발주량"]
        
        def on_edit_5():
            edits = st.session_state["editor_v5"]["edited_rows"]
            for r_idx_str, change in edits.items():
                if "추가발주수량" in change:
                    r_key = to_order.iloc[int(r_idx_str)]['unique_key']
                    st.session_state.extra_order_dict[r_key] = int(change["추가발주수량"])
            st.rerun()

        st.data_editor(to_order[disp_final], use_container_width=True, key="editor_v5", on_change=on_edit_5, hide_index=True)

        b1, b2 = st.columns(2)
        if b1.button("💾 구글 시트에 최종 발주 기록 저장", use_container_width=True):
            to_save = to_order[to_order['최종발주량'] > 0].copy()
            if not to_save.empty:
                if save_history_to_gsheet(to_save[[item, option, '최종발주량']], log_type="발주"):
                    st.success("✅ 저장 성공!")
            else: st.warning("발주 항목 없음")

        csv_v5 = to_order.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
        b2.download_button("📥 엑셀(CSV) 다운로드", data=csv_v5, file_name=f"발주서_{datetime.now().strftime('%m%d')}.csv", use_container_width=True)

        # --- [6단계: 기록 통합 조회] ---
        st.divider()
        st.subheader("📜 6단계: 제작 상품 입고 및 발주 히스토리")
        c6_1, c6_2 = st.columns(2)
        start_d = c6_1.date_input("📅 조회 시작일", datetime.now() - timedelta(days=7))
        end_d = c6_2.date_input("📅 조회 종료일", datetime.now())

        if not past_hist.empty:
            past_hist['날짜_dt'] = pd.to_datetime(past_hist['저장시간']).dt.date
            df_h = past_hist[(past_hist['날짜_dt'] >= start_d) & (past_hist['날짜_dt'] <= end_d)].copy()
            
            t_in, t_out = st.tabs(["📥 입고 완료 내역", "📤 발주 진행 내역"])
            with t_in:
                in_df = df_h[df_h['구분'] == "입고"]
                st.dataframe(in_df[['저장시간', '상품명', '옵션', '수량']], use_container_width=True, hide_index=True)
                if not in_df.empty:
                    st.table(in_df.groupby(['상품명', '옵션'])['수량'].sum().reset_index())
            with t_out:
                out_df = df_h[df_h['구분'] == "발주"]
                st.dataframe(out_df[['저장시간', '상품명', '옵션', '수량']], use_container_width=True, hide_index=True)
                if not out_df.empty:
                    st.table(out_df.groupby(['상품명', '옵션'])['수량'].sum().reset_index())

# --- [탭 2: 동대문 상품 관리 (원본 로직 유지)] ---
with tab2:
    st.subheader("🌙 동대문 사입 가중율 관리")
    d_file = st.file_uploader("동대문 리스트 업로드", type=['xlsx', 'xls', 'csv'], key="d_up")
    if d_file:
        if "df_dong" not in st.session_state:
            df = pd.read_excel(d_file) if d_file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(d_file)
            df['판매수량'] = (safe_num(df.get('정상재고', 0)) - safe_num(df.get('가용재고', 0))).clip(lower=0)
            df['가중율'] = df['판매수량'].apply(lambda n: 2.0 if n >= 10 else (1.5 if n >= 6 else (1.2 if n >= 3 else 1.0)))
            df['발주수량'] = (df['판매수량'] * df['가중율']).astype(int)
            df['선택'] = False
            st.session_state.df_dong = df

        if "df_dong" in st.session_state:
            c1, c2 = st.columns([1, 1])
            add_v = c1.number_input("➕ 더할 수량", value=1)
            if c2.button("🚀 선택 상품 수량 더하기"):
                for i, row in ed_dong.iterrows():
                    if row['선택']: st.session_state.df_dong.at[i, '발주수량'] += add_v
                st.rerun()

            ed_dong = st.data_editor(st.session_state.df_dong, use_container_width=True, key="ed_dong", 
                                     column_config={"선택": st.column_config.CheckboxColumn("선택", default=False)}, hide_index=True)
