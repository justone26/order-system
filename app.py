import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import io

# --- [기본 함수 설정] ---
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
        data = hist_sheet.get_all_records()
        return pd.DataFrame(data) if data else pd.DataFrame(columns=["저장시간", "구분", "상품명", "옵션", "수량"])
    except: return pd.DataFrame(columns=["저장시간", "구분", "상품명", "옵션", "수량"])

def find_idx(cols, target_keywords):
    for keyword in target_keywords:
        for i, col in enumerate(cols):
            if keyword in str(col): return i
    return 0

def safe_num(val):
    res = pd.to_numeric(val, errors='coerce')
    if isinstance(res, pd.Series): return res.fillna(0)
    return 0 if pd.isna(res) else res

# --- [초기 설정] ---
st.set_page_config(layout="wide", page_title="저스트원 재고관리")
st.title("🏭 저스트원 통합 재고 관리 시스템")

if "extra_order_dict" not in st.session_state: st.session_state.extra_order_dict = {}
if 'analyzed' not in st.session_state: st.session_state.analyzed = False

tab1, tab2 = st.tabs(["✂️ 제작 상품 관리", "🌙 동대문 상품 관리"])

# --- [탭 1: 제작 상품 관리] ---
with tab1:
    uploaded_file = st.file_uploader("엑셀 파일을 올려주세요", type=['xlsx', 'xls', 'csv'], key="tab1_upload")
    
    # 이전 데이터 로드 기능 유지
    if st.button("📂 이전 데이터 로드", use_container_width=True):
        try:
            sheet = get_sheet().sheet1
            gs_df = pd.DataFrame(sheet.get_all_records())
            if not gs_df.empty:
                st.session_state.df_raw = gs_df.copy()
                st.session_state.analyzed = False
                st.success("데이터 로드 완료! 매핑을 확인해주세요.")
                st.rerun()
        except: st.error("연결 실패")

    if uploaded_file:
        if 'df_raw' not in st.session_state or st.session_state.get('last_filename') != uploaded_file.name:
            df_new = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
            df_new.columns = df_new.columns.str.strip()
            # 리오더 수량 매핑 로직 유지
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
            st.session_state.last_filename = uploaded_file.name
            st.session_state.analyzed = False
            st.rerun()

    if st.session_state.get('df_raw') is not None:
        df_curr = st.session_state.df_raw
        cols = df_curr.columns.tolist()
        
        st.subheader("⚙️ 매핑 설정")
        col_left, col_right = st.columns(2)
        with col_left:
            sold_out = st.selectbox("품절 여부", cols, index=find_idx(cols, ['품절']))
            vendor = st.selectbox("공급처", cols, index=find_idx(cols, ['공급처']))
            v_item = st.selectbox("공급처 상품명", cols, index=find_idx(cols, ['공급처상품명']))
            item = st.selectbox("상품명", cols, index=find_idx(cols, ['상품명']))
            option = st.selectbox("옵션", cols, index=find_idx(cols, ['옵션']))
        with col_right:
            reg_date = st.selectbox("등록일", cols, index=find_idx(cols, ['등록일']))
            stock = st.selectbox("정상재고", cols, index=find_idx(cols, ['정상재고']))
            avail = st.selectbox("가용재고", cols, index=find_idx(cols, ['가용재고']))
            t3day = st.selectbox("3일 발주합계", cols, index=find_idx(cols, ['3일']))
            t7day = st.selectbox("7일 발주합계", cols, index=find_idx(cols, ['7일', '1주']))

        st.subheader("🚀 분석 설정")
        l1, l2 = st.columns(2)
        lt = l1.number_input("리드타임 (일)", value=7)
        ss = l2.number_input("안전재고 (일 수)", value=3)
        if st.button("📊 분석 실행", use_container_width=True):
            st.session_state.analyzed = True
            st.rerun()
            
    if st.session_state.get('analyzed'):
        st.divider()
        st.subheader("📊 4단계: 데이터 편집 및 재고 관리")
        df_work = st.session_state.df_raw.copy()
        
        # 검색 및 필터 UI
        f_c1, f_c2, f_c3 = st.columns([2, 1, 1])
        search_q = f_c1.text_input("🔍 상품명 검색", key="search_q_t1")
        filter_m = f_c2.selectbox("품절 필터", ["전체보기", "정상만", "품절만"], index=1)
        hist_date_in = f_c3.date_input("🗓️ 입고 기록 날짜", datetime.now())

        # 계산 로직
        def simple_key(n): return str(n).strip().replace(" ", "").upper() if not pd.isna(n) else ""
        df_work['unique_key'] = df_work[item].apply(simple_key) + df_work[option].apply(simple_key)
        
        # 수치 계산
        v7 = safe_num(df_work[t7day])
        v3 = safe_num(df_work[t3day])
        df_work['일판매량'] = (v7 / 7 if v7.sum() > 0 else v3 / 3).round(0).astype(int)
        df_work['권장발주량'] = ((df_work['일판매량'] * (lt + ss)) - (safe_num(df_work[avail]) + safe_num(df_work.get('리오더 수량', 0)))).clip(lower=0).astype(int)

        # 필터 적용
        if filter_m == "정상만": df_work = df_work[~df_work[sold_out].astype(str).str.contains('품절', na=False)]
        elif filter_m == "품절만": df_work = df_work[df_work[sold_out].astype(str).str.contains('품절', na=False)]
        if search_q: df_work = df_work[df_work[item].astype(str).str.contains(search_q, case=False, na=False)]

        # 데이터 편집기
        if not df_work.empty:
            display_cols = [sold_out, vendor, item, option, stock, avail, "리오더 수량", "리오더입고수량", t3day, "일판매량", "권장발주량"]
            df_view = df_work[[c for c in display_cols if c in df_work.columns or c == "리오더입고수량"]].copy()
            if "리오더입고수량" not in df_view.columns: df_view["리오더입고수량"] = 0
            
            def on_edit_t1():
                changes = st.session_state["editor_t1"]["edited_rows"]
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

            st.data_editor(df_view, use_container_width=True, key="editor_t1", on_change=on_edit_t1, hide_index=True)

        # 5단계: 최종 요약 (수정된 로직 유지)
        st.divider()
        st.subheader("📋 5단계: 최종 발주 리스트 요약")
        to_order = df_work.copy()
        to_order['추가발주수량'] = to_order['unique_key'].map(st.session_state.extra_order_dict).fillna(0).astype(int)
        to_order['최종발주량'] = to_order['권장발주량'] + to_order['추가발주수량']
        
        # 상태 판별 로직 복구
        def get_stat(r):
            total = safe_num(r[avail]) + safe_num(r.get('리오더 수량', 0))
            if r['일판매량'] > 0:
                if total < (r['일판매량'] * 3): return "🚨 긴급"
                if total < (r['일판매량'] * 5): return "⚠️ 주의"
            return "✅ 정상"
        to_order['상태'] = to_order.apply(get_stat, axis=1)
        
        # 최종 리스트 출력
        f_view = ["상태", item, option, vendor, avail, "리오더 수량", "추가발주수량", "권장발주량", "최종발주량"]
        df_final = to_order[[c for c in f_view if c in to_order.columns]].copy()
        
        def on_edit_f():
            for r_idx, change in st.session_state["editor_f"]["edited_rows"].items():
                if "추가발주수량" in change:
                    st.session_state.extra_order_dict[to_order.iloc[int(r_idx)]['unique_key']] = int(change["추가발주수량"])
            st.rerun()
        
        st.data_editor(df_final, use_container_width=True, key="editor_f", on_change=on_edit_f, hide_index=True)

# --- [탭 2: 동대문 상품 관리] ---
with tab2:
    st.subheader("🌙 동대문 사입 및 미납 관리")
    dong_file = st.file_uploader("동대문 주문 리스트 업로드", type=['xlsx', 'xls', 'csv'], key="dong_upload_final")

    if dong_file:
        if "last_dong_name" not in st.session_state or st.session_state.last_dong_name != dong_file.name:
            df = pd.read_excel(dong_file) if dong_file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(dong_file)
            df.columns = df.columns.str.strip()
            
            # 사장님의 가중율 로직 그대로 복구
            req = ['선택', '품절', '상품명', '공급처', '공급처상품명', '정상재고', '가용재고', '판매수량', '발주수량', '가중율', '3일판매']
            for c in req: 
                if c not in df.columns: df[c] = 0 if c not in ['선택', '품절', '상품명', '공급처', '공급처상품명'] else ""
            
            df['판매수량'] = (safe_num(df['정상재고']) - safe_num(df['가용재고'])).clip(lower=0)
            df['가중율'] = df['판매수량'].apply(lambda n: 2.0 if n >= 10 else (1.5 if n >= 6 else (1.2 if n >= 3 else 1.0)))
            df['발주수량'] = (df['판매수량'] * df['가중율']).astype(int)
            df['선택'] = False
            
            st.session_state.df_dong_current = df[req]
            st.session_state.last_dong_name = dong_file.name

        if "df_dong_current" in st.session_state:
            search_dong = st.text_input("🔍 상품명 검색 (사입)")
            df_dong_disp = st.session_state.df_dong_current.copy()
            if search_dong:
                df_dong_disp = df_dong_disp[df_dong_disp['상품명'].astype(str).str.contains(search_dong, case=False)]
            
            # 에디터 및 체크박스 기능
            edited_dong = st.data_editor(
                df_dong_disp, 
                use_container_width=True, 
                key="dong_editor_final",
                column_config={"선택": st.column_config.CheckboxColumn("선택", default=False)},
                hide_index=True
            )
            
            # 수량 더하기 버튼 로직 복구
            c1, c2, c3 = st.columns([1, 1, 1])
            add_qty = c1.number_input("➕ 추가 수량", value=1, min_value=1)
            if c2.button("🚀 선택 상품 수량 더하기", use_container_width=True):
                for i, row in edited_dong.iterrows():
                    if row['선택']:
                        st.session_state.df_dong_current.at[i, '발주수량'] += add_qty
                st.rerun()
            
            csv_dong = edited_dong.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            c3.download_button("📥 사입 리스트 다운로드", data=csv_dong, file_name=f"동대문사입_{datetime.now().strftime('%m%d')}.csv", use_container_width=True)
