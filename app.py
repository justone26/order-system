import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
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
    except:
        return None

def make_match_key(name, opt):
    """상품명과 옵션으로 고유 키 생성 (공백 제거 및 대문자화)"""
    return str(name).replace(" ", "").upper() + str(opt).replace(" ", "").upper()

def save_reorder_data(new_df):
    """[수정본] 기존 시트 데이터를 유지하며 현재 데이터만 누적/업데이트 저장"""
    try:
        spreadsheet = get_sheet()
        sheet = spreadsheet.sheet1
        
        # 1. 기존 구글 시트 데이터 로드
        gs_raw = sheet.get_all_records()
        gs_df = pd.DataFrame(gs_raw)
        
        # 2. 저장할 데이터에 비교 키 생성
        new_df['k_tmp'] = new_df.apply(lambda r: make_match_key(r['상품명'], r['옵션']), axis=1)
        
        if gs_df.empty:
            final_df = new_df.drop(columns=['k_tmp'])
        else:
            # 3. 기존 데이터에도 키 생성
            gs_df['k_tmp'] = gs_df.apply(lambda r: make_match_key(r['상품명'], r['옵션']), axis=1)
            
            # 4. [핵심] 현재 업로드된 상품이 아닌 데이터(타 업체 데이터)만 추출
            others_df = gs_df[~gs_df['k_tmp'].isin(new_df['k_tmp'])].copy()
            
            # 5. 기존 타 업체 데이터 + 현재 새 데이터를 합침
            final_df = pd.concat([others_df, new_df], ignore_index=True)
            final_df = final_df.drop(columns=['k_tmp'])

        # 6. 시트 갱신 (전체 덮어쓰기지만 데이터는 합쳐진 상태)
        sheet.clear()
        sheet.update([final_df.columns.values.tolist()] + final_df.fillna(0).values.tolist())
        return True
    except Exception as e:
        st.error(f"시트 저장 실패: {e}")
        return False

def save_history_to_gsheet(df, log_type="발주"):
    try:
        spreadsheet = get_sheet()
        try:
            hist_sheet = spreadsheet.worksheet("history")
        except:
            hist_sheet = spreadsheet.add_worksheet(title="history", rows="1000", cols="20")
            hist_sheet.append_row(["저장시간", "구분", "상품명", "옵션", "수량"])
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        rows_to_add = [[now_str, log_type] + [str(x) for x in row] for row in df.values.tolist()]
        hist_sheet.append_rows(rows_to_add)
        return True
    except:
        return False

def load_history_from_gsheet():
    try:
        spreadsheet = get_sheet()
        hist_sheet = spreadsheet.worksheet("history")
        df = pd.DataFrame(hist_sheet.get_all_records())
        df.columns = df.columns.str.strip()
        return df
    except:
        return pd.DataFrame()

def find_idx(cols, target_keywords):
    for keyword in target_keywords:
        for i, col in enumerate(cols):
            if keyword in str(col): return i
    return 0

def safe_num(val):
    res = pd.to_numeric(val, errors='coerce')
    return res.fillna(0) if hasattr(res, 'fillna') else (0 if pd.isna(res) else res)

# --- [2. 앱 설정 및 세션 초기화] ---
if "extra_order_dict" not in st.session_state:
    st.session_state.extra_order_dict = {}
if 'analyzed' not in st.session_state: 
    st.session_state.analyzed = False

st.set_page_config(layout="wide", page_title="재고 관리 시스템")
st.title("📦 통합 재고 관리 시스템")

tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 관리"])

with tab1:
    st.subheader("📁 데이터 업로드 (제작상품)")
    if st.button("🔄 제작상품 데이터 초기화"):
        for key in ['df_raw', 'analyzed', 'last_filename', 'extra_order_dict']:
            if key in st.session_state: del st.session_state[key]
        st.rerun()

    uploaded_file = st.file_uploader("엑셀/CSV 파일을 선택하세요", type=['xlsx', 'xls', 'csv'], key="prod_upload")

    if uploaded_file is not None:
        if 'df_raw' not in st.session_state or st.session_state.get('last_filename') != uploaded_file.name:
            df_new = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
            df_new.columns = df_new.columns.str.strip()
            df_new = df_new.loc[:, ~df_new.columns.duplicated()] 
            
            try:
                sheet_obj = get_sheet().sheet1
                gs_data = pd.DataFrame(sheet_obj.get_all_records())
                if not gs_data.empty and '리오더 수량' in gs_data.columns:
                    t_item = next((c for c in df_new.columns if '상품명' in c), df_new.columns[0])
                    t_opt = next((c for c in df_new.columns if '옵션' in c), df_new.columns[1])
                    
                    df_new['k_tmp'] = df_new.apply(lambda r: make_match_key(r[t_item], r[t_opt]), axis=1)
                    gs_data['k_tmp'] = gs_data.apply(lambda r: make_match_key(r['상품명'], r['옵션']), axis=1)
                    
                    reorder_map = gs_data.set_index('k_tmp')['리오더 수량'].to_dict()
                    df_new['리오더 수량'] = df_new['k_tmp'].map(reorder_map).fillna(0).astype(int)
                    df_new.drop(columns=['k_tmp'], inplace=True)
                else: 
                    df_new['리오더 수량'] = 0
            except: 
                df_new['리오더 수량'] = 0
                
            st.session_state.df_raw = df_new
            st.session_state.last_filename = uploaded_file.name
            st.rerun()

    if st.session_state.get('df_raw') is not None:
        df_curr = st.session_state.df_raw
        cols = df_curr.columns.tolist()

        # --- 1단계: 매핑 설정 ---
        st.subheader("⚙️ 1단계: 매핑 설정")
        c1, c2 = st.columns(2)
        sold_out = c1.selectbox("품절 여부", cols, index=find_idx(cols, ['품절']))
        vendor = c1.selectbox("공급처", cols, index=find_idx(cols, ['공급처']))
        item = c1.selectbox("상품명", cols, index=find_idx(cols, ['상품명']))
        option = c1.selectbox("옵션", cols, index=find_idx(cols, ['옵션']))
        avail = c2.selectbox("가용재고", cols, index=find_idx(cols, ['가용재고']))
        t3day = c2.selectbox("3일 발주합계", cols, index=find_idx(cols, ['3일']))
        t7day = c2.selectbox("7일 발주합계", cols, index=find_idx(cols, ['7일', '1주']))

        # --- 2~3단계: 분석 설정 ---
        st.subheader("🚀 2~3단계: 분석 설정")
        l1, l2 = st.columns(2)
        lt = l1.number_input("리드타임 (일)", value=10)
        ss = l2.number_input("안전재고 (일 수)", value=7)
        if st.button("📊 분석 실행", use_container_width=True):
            st.session_state.analyzed = True
            st.rerun()

    if st.session_state.get('analyzed'):
        # --- 4단계: 데이터 편집 및 재고 관리 ---
        st.divider()
        st.subheader("📊 4단계: 데이터 편집 및 재고 관리")
        df_work = st.session_state.df_raw.copy()
        
        f_c1, f_c2, f_c3 = st.columns([2, 1, 1])
        search_q = f_c1.text_input("🔍 상품명 검색")
        filter_m = f_c2.selectbox("품절 필터", ["전체보기", "정상만", "품절만"], index=1)
        hist_date_4 = f_c3.date_input("🗓️ 입고 기록 날짜", datetime.now())

        df_work['unique_key'] = df_work.apply(lambda r: make_match_key(r[item], r[option]), axis=1)

        # 수치 계산
        v7, v3 = safe_num(df_work[t7day]), safe_num(df_work[t3day])
        df_work['일판매량'] = (v7 / 7 if v7.sum() > 0 else v3 / 3).round(1)
        df_work['권장발주량'] = ((df_work['일판매량'] * (lt + ss)) - (safe_num(df_work[avail]) + safe_num(df_work['리오더 수량']))).clip(lower=0).astype(int)
        df_work['리오더입고수량'] = 0

        # 필터링
        if filter_m == "정상만": df_work = df_work[~df_work[sold_out].astype(str).str.contains('품절', na=False)]
        elif filter_m == "품절만": df_work = df_work[df_work[sold_out].astype(str).str.contains('품절', na=False)]
        if search_q: df_work = df_work[df_work[item].astype(str).str.contains(search_q, case=False, na=False)]

        disp4 = [vendor, item, option, avail, "리오더 수량", "리오더입고수량", t3day, "일판매량", "권장발주량"]
        edited4 = st.data_editor(df_work[disp4], use_container_width=True, hide_index=True, key="editor4")

        if st.button("💾 리오더 현황 및 입고 저장", use_container_width=True):
            for idx, row in edited4.iterrows():
                o_idx = df_work.index[idx]
                st.session_state.df_raw.at[o_idx, "리오더 수량"] = int(row["리오더 수량"])
                in_qty = int(row["리오더입고수량"])
                if in_qty > 0:
                    st.session_state.df_raw.at[o_idx, "리오더 수량"] = max(0, int(st.session_state.df_raw.at[o_idx, "리오더 수량"]) - in_qty)
                    save_history_to_gsheet(pd.DataFrame([[row[item], row[option], in_qty]], columns=['상품명', '옵션', '수량']), "입고")
            
            # [누적 저장 호출]
            save_df = st.session_state.df_raw[[item, option, '리오더 수량']].rename(columns={item:'상품명', option:'옵션'})
            if save_reorder_data(save_df):
                st.success("✅ 구글 시트에 누적 저장되었습니다!")
                st.rerun()

        # --- 5단계: 최종 발주 요약 ---
        st.divider()
        st.subheader("📋 5단계: 최종 발주 요약")
        to_order = df_work.copy()
        to_order['추가발주수량'] = to_order['unique_key'].map(st.session_state.extra_order_dict).fillna(0).astype(int)
        to_order['최종발주량'] = to_order['권장발주량'] + to_order['추가발주수량']

        def get_stat(r):
            total = safe_num(r[avail]) + safe_num(r['리오더 수량'])
            if r['일판매량'] > 0:
                if total < (r['일판매량'] * 3): return "🚨 긴급"
                if total < (r['일판매량'] * 5): return "⚠️ 주의"
            return "✅ 정상"
        to_order['상태'] = to_order.apply(get_stat, axis=1)

        disp5 = ["상태", item, option, vendor, avail, "리오더 수량", "추가발주수량", "권장발주량", "최종발주량"]
        edited5 = st.data_editor(to_order[disp5], use_container_width=True, hide_index=True, key="editor5")

        for idx, row in edited5.iterrows():
            k = to_order.iloc[idx]['unique_key']
            st.session_state.extra_order_dict[k] = int(row["추가발주수량"])

        if st.button("📄 발주 기록 저장", use_container_width=True):
            final_orders = edited5[edited5['최종발주량'] > 0]
            if not final_orders.empty:
                save_history_to_gsheet(final_orders[[item, option, '최종발주량']].rename(columns={item:'상품명', option:'옵션', '최종발주량':'수량'}), "발주")
                st.success("발주 기록 완료")

# --- [6단계: 과거 데이터 조회 탭] ---
    st.divider()
    st.subheader("📜 6단계: 히스토리 조회")
    hist_all = load_history_from_gsheet()
    if not hist_all.empty:
        st.dataframe(hist_all, use_container_width=True, hide_index=True)

# --- [🌙 탭 2: 동대문 사입 관리] ---
with tab2:
    st.subheader("🌙 동대문 사입 및 미납 관리")
    dong_file = st.file_uploader("동대문 주문 리스트 업로드", type=['xlsx', 'csv'], key="dong_tab_upload")
    if dong_file:
        df_dong = pd.read_excel(dong_file)
        st.dataframe(df_dong, use_container_width=True)
        csv_dong = df_dong.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
        st.download_button("📥 엑셀 다운로드", csv_dong, "사입리스트.csv")
