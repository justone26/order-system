import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta, timezone
import time

# 한국 시간(KST) 설정
KST = timezone(timedelta(hours=9))

# --- [1. 구글 시트 및 유틸 함수] ---
def get_sheet():
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        spreadsheet_key = "1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg"
        return client.open_by_key(spreadsheet_key)
    except Exception as e:
        st.error(f"🚨 구글 시트 연결 실패: {e}")
        return None

def make_match_key(item, opt):
    return str(item).strip().replace(" ", "").upper() + str(opt).strip().replace(" ", "").upper()

def find_idx(cols, target_keywords):
    for keyword in target_keywords:
        for i, col in enumerate(cols):
            if keyword in str(col): return i
    return 0

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
        
        if not gs_df.empty:
            gs_df['리오더 수량'] = pd.to_numeric(gs_df['리오더 수량'], errors='coerce').fillna(0)

        for _, row in new_work_df.iterrows():
            target_key = row['match_key']
            if target_key in gs_df['match_key'].values:
                gs_df.loc[gs_df['match_key'] == target_key, '리오더 수량'] += row['리오더 수량']
            else:
                new_item = pd.DataFrame([{'상품명': row['상품명'], '옵션': row['옵션'], '리오더 수량': row['리오더 수량']}])
                gs_df = pd.concat([gs_df, new_item], ignore_index=True)

        final_df = gs_df.drop(columns=['match_key'], errors='ignore').fillna(0)
        final_df = final_df.drop_duplicates(subset=['상품명', '옵션'], keep='last')
        sheet.clear()
        sheet.update([final_df.columns.values.tolist()] + final_df.values.tolist())
        return True
    except Exception as e:
        st.error(f"⚠️ 데이터 누적 저장 중 오류 발생: {e}")
        return False

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
        if rows_to_add:
            hist_sheet.append_rows(rows_to_add)
            return True
        return False
    except Exception as e:
        st.error(f"히스토리 저장 실패: {e}")
        return False

# --- [2. 앱 초기 설정] ---
st.set_page_config(layout="wide", page_title="저스트원 재고관리")
st.title("🏭 저스트원 통합 재고 관리 시스템")

if "extra_order_dict" not in st.session_state: st.session_state.extra_order_dict = {}
if "add_order_dict" not in st.session_state: st.session_state.add_order_dict = {}
if 'analyzed' not in st.session_state: st.session_state.analyzed = False

tab1, tab2 = st.tabs(["✂️ 제작 상품 관리", "🌙 동대문 상품 관리"])

# --- [탭 1: 제작 상품 관리] ---
with tab1:
    uploaded_file = st.file_uploader("엑셀 파일을 올려주세요", type=['xlsx', 'xls', 'csv'], key="t1_up")
    
    if st.button("📂 구글 시트 데이터 로드", use_container_width=True):
        spreadsheet = get_sheet()
        if spreadsheet:
            try:
                with st.spinner('📡 시트 데이터를 불러오는 중...'):
                    sheet = spreadsheet.get_worksheet(0)
                    raw_data = sheet.get_all_values()
                    if len(raw_data) > 1:
                        header = [str(h).strip() for h in raw_data[0]]
                        df_tmp = pd.DataFrame(raw_data[1:], columns=header)
                        st.session_state.df_raw = df_tmp.copy()
                        st.session_state.analyzed = True 
                        st.success(f"✅ {len(df_tmp)}건 로드 완료!")
                        time.sleep(1)
                        st.rerun()
            except Exception as e: st.error(f"❌ 오류: {e}")

    if uploaded_file:
        if 'df_raw' not in st.session_state or st.session_state.get('last_fn') != uploaded_file.name:
            df_new = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
            df_new.columns = df_new.columns.str.strip()
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
            st.rerun()

    if st.session_state.get('df_raw') is not None:
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
            sel_reg_date = st.selectbox("등록일", cols, index=find_idx(cols, ['등록일']))
            sel_stock = st.selectbox("정상재고", cols, index=find_idx(cols, ['정상재고']))
            sel_avail = st.selectbox("가용재고", cols, index=find_idx(cols, ['가용재고']))
            sel_t3day = st.selectbox("3일 발주합계", cols, index=find_idx(cols, ['3일']))
            sel_t7day = st.selectbox("7일 발주합계", cols, index=find_idx(cols, ['7일', '1주']))

        lt = st.number_input("리드타임 (일)", value=7)
        ss = st.number_input("안전재고 (일 수)", value=3)

        # --- [4단계: 사장님 로직 그대로 유지] ---
        st.divider()
        st.subheader("📊 4단계: 데이터 편집 및 재고 관리")
        
        df_work = df_curr.copy()
        # 매핑된 컬럼명으로 숫자 변환
        for c in [sel_stock, sel_avail, "리오더 수량", sel_t7day, sel_t3day]:
            if c in df_work.columns:
                df_work[c] = pd.to_numeric(df_work[c], errors='coerce').fillna(0).astype(int)

        # 사장님의 계산 공식
        v7 = df_work[sel_t7day]; v3 = df_work[sel_t3day]
        df_work['일판매량'] = (v7 / 7 if v7.sum() > 0 else v3 / 3).round(0).astype(int)
        df_work['권장발주량'] = ((df_work['일판매량'] * (lt + ss)) - (df_work[sel_avail] + df_work['리오더 수량'])).clip(lower=0).astype(int)
        
        # 필터 및 검색
        f_c1, f_c2, f_c3 = st.columns([2, 1, 1])
        search_q = f_c1.text_input("🔍 상품명 검색", key="search_v4")
        filter_m = f_c2.selectbox("품절 필터", ["전체보기", "정상만", "품절만"], index=1, key="filter_v4")
        hist_date_4 = f_c3.date_input("🗓️ 입고 매핑 날짜", datetime.now(KST).date(), key="date_v4")

        if filter_m == "정상만": df_work = df_work[~df_work[sel_sold_out].astype(str).str.contains('품절', na=False)]
        elif filter_m == "품절만": df_work = df_work[df_work[sel_sold_out].astype(str).str.contains('품절', na=False)]
        if search_q: df_work = df_work[df_work[sel_item].astype(str).str.contains(search_q, case=False, na=False)]

        # 표시용 컬럼 정리
        df_display = df_work.rename(columns={sel_sold_out: "품절", sel_vendor: "공급쳐", sel_v_item: "공급쳐 상품명", sel_item: "상품명", sel_option: "옵션", sel_stock: "정상재고", sel_avail: "가용재고"})
        actual_final_cols = ["품절", "공급쳐", "상품명", "옵션", "공급쳐 상품명", "정상재고", "가용재고", "리오더 수량", "일판매량", "권장발주량"]
        
        with st.form("form_step_4"):
            edited_v4 = st.data_editor(df_display[[c for c in actual_final_cols if c in df_display.columns]], use_container_width=True, key="editor_v4", hide_index=True)
            if st.form_submit_button("💾 입고량 반영 및 저장", use_container_width=True, type="primary"):
                edits = st.session_state["editor_v4"].get("edited_rows", {})
                if edits:
                    for r_idx_str, change in edits.items():
                        orig_idx = df_work.index[int(r_idx_str)]
                        if "리오더 수량" in change: st.session_state.df_raw.at[orig_idx, "리오더 수량"] = int(change["리오더 수량"])
                    save_reorder_data(st.session_state.df_raw[[sel_item, sel_option, '리오더 수량']].rename(columns={sel_item:'상품명', sel_option:'옵션'}))
                    st.success("✅ 저장 완료!")
                    st.rerun()

        # --- [5단계: 사장님 로직 그대로 유지] ---
        st.divider()
        st.subheader("📋 5단계: 최종 발주 리스트 요약")
        df_5 = st.session_state.df_raw.copy()
        for c in [sel_avail, '리오더 수량', sel_t7day, sel_t3day]:
            df_5[c] = pd.to_numeric(df_5[c], errors='coerce').fillna(0).astype(int)

        df_5['일판매량'] = (df_5[sel_t7day] / 7 if df_5[sel_t7day].sum() > 0 else df_5[sel_t3day] / 3).round(0).astype(int)
        df_5['권장발주량'] = ((df_5['일판매량'] * (lt + ss)) - (df_5[sel_avail] + df_5['리오더 수량'])).clip(lower=0).astype(int)
        df_5['추가발주수량'] = df_5.index.map(st.session_state.add_order_dict).fillna(0).astype(int)

        def get_status(r):
            s_sum = r[sel_avail] + r['리오더 수량']; d = r['일판매량']
            if d > 0:
                if s_sum < (d * 3): return "🚨 긴급"
                if s_sum < (d * 5): return "⚠️ 주의"
            return "✅ 정상"
        df_5['상태'] = df_5.apply(get_status, axis=1)

        c5_1, c5_2 = st.columns([2, 1])
        s_filter = c5_1.selectbox("🎯 상태 필터", ["🚨긴급 + ⚠️주의 우선", "🚨 긴급만 보기", "✅ 전체보기"], index=0, key="filter_v5")
        if s_filter == "🚨긴급 + ⚠️주의 우선": df_5 = df_5[df_5['상태'].isin(["🚨 긴급", "⚠️ 주의"]) | (df_5['권장발주량'] > 0)]
        elif s_filter == "🚨 긴급만 보기": df_5 = df_5[df_5['상태'] == "🚨 긴급"]

        df_display_5 = df_5.rename(columns={sel_item: "상품명", sel_option: "옵션", sel_v_item: "공급쳐상품명", sel_avail: "가용재고"})
        with st.form("form_step_5"):
            edited_v5 = st.data_editor(df_display_5[["상태", "상품명", "옵션", "공급쳐상품명", "가용재고", "리오더 수량", "추가발주수량", "권장발주량"]], use_container_width=True, key="editor_v5", hide_index=True)
            if st.form_submit_button("✅ 수량 확정 (리오더 수량 합산)", use_container_width=True):
                edits = st.session_state["editor_v5"].get("edited_rows", {})
                for r_idx_str, change in edits.items():
                    orig_idx = df_5.index[int(r_idx_str)]
                    if "추가발주수량" in change:
                        st.session_state.df_raw.at[orig_idx, "리오더 수량"] += int(change["추가발주수량"])
                        st.session_state.add_order_dict[orig_idx] = int(change["추가발주수량"])
                st.success("✅ 갱신 완료!")
                st.rerun()

        st.write("---")
        b1, b2 = st.columns(2)
        if b1.button("💾 구글 시트에 최종 발주 기록 저장", use_container_width=True):
            order_ready = df_5[(df_5['권장발주량'] > 0) | (df_5['추가발주수량'] > 0)].copy()
            if not order_ready.empty:
                save_data = order_ready[[sel_item, sel_option, '권장발주량']] # 예시 저장 컬럼
                if save_history_to_gsheet(save_data, log_type="발주"): st.success("✅ 저장 성공!")
        
        csv_final = df_display_5.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
        b2.download_button(label="📥 엑셀 다운로드", data=csv_final, file_name="발주서.csv", use_container_width=True)

        # --- [6단계: 히스토리] ---
        st.divider()
        st.subheader("📜 6단계: 전체 히스토리 내역")
        df_hist = load_history_from_gsheet()
        if not df_hist.empty: st.dataframe(df_hist.sort_index(ascending=False), use_container_width=True)

# --- [탭 2: 동대문 사입 관리 - 사장님 로직 그대로] ---
with tab2:
    st.subheader("🌙 동대문 사입 및 미납 관리")
    dong_file = st.file_uploader("동대문 주문 리스트 업로드", type=['xlsx', 'xls', 'csv'], key="dong_up")
    if dong_file:
        df_dong = pd.read_excel(dong_file) if dong_file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(dong_file)
        df_dong.columns = df_dong.columns.str.strip()
        # 사장님의 가중치 계산 로직
        df_dong['판매수량'] = (pd.to_numeric(df_dong['정상재고'], errors='coerce').fillna(0) - pd.to_numeric(df_dong['가용재고'], errors='coerce').fillna(0)).clip(lower=0)
        df_dong['발주수량'] = (df_dong['판매수량'] * df_dong['판매수량'].apply(lambda n: 2.0 if n >= 10 else 1.0)).astype(int)
        st.data_editor(df_dong, use_container_width=True, hide_index=True)
