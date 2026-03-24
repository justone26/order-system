import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import io

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
    except:
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
        data = hist_sheet.get_all_records()
        return pd.DataFrame(data) if data else pd.DataFrame(columns=["저장시간", "구분", "상품명", "옵션", "수량"])
    except:
        return pd.DataFrame(columns=["저장시간", "구분", "상품명", "옵션", "수량"])

def find_idx(cols, target_keywords):
    for keyword in target_keywords:
        for i, col in enumerate(cols):
            if keyword in str(col): return i
    return 0

def safe_num(val):
    res = pd.to_numeric(val, errors='coerce')
    if isinstance(res, pd.Series):
        return res.fillna(0)
    return 0 if pd.isna(res) else res

st.set_page_config(layout="wide", page_title="저스트원 재고관리")
st.title("🏭 제작 상품 재고 관리 시스템")

if "extra_order_dict" not in st.session_state:
    st.session_state.extra_order_dict = {}
if 'analyzed' not in st.session_state:
    st.session_state.analyzed = False

tab1, tab2 = st.tabs(["📊 제작 상품 관리", "📜 히스토리 조회"])

with tab1:
    uploaded_file = st.file_uploader("엑셀 파일을 올려주세요", type=['xlsx', 'xls', 'csv'])
    c_btn, _ = st.columns([1, 4])
    if c_btn.button("📂 이전 데이터 로드", width='stretch'):
        try:
            sheet = get_sheet().sheet1
            gs_df = pd.DataFrame(sheet.get_all_records())
            if not gs_df.empty:
                st.session_state.df_raw = gs_df.rename(columns={'상품명': '기존상품명', '옵션': '기존옵션'})
                st.success("로드 완료")
            else:
                st.warning("데이터 없음")
        except:
            st.error("연결 오류")

    if uploaded_file:
        if 'df_raw' not in st.session_state or st.session_state.get('last_filename') != uploaded_file.name:
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
        lt = l1.number_input("리드타임 (일)", value=10)
        ss = l2.number_input("안전재고 (일 수)", value=7)
        if st.button("📊 분석 실행", width='stretch'):
            st.session_state.analyzed = True
            st.rerun()

    if st.session_state.get('analyzed'):
        st.divider()
        df_all = st.session_state.df_raw.copy()
        for c in [v_item, item, option]:
            df_all[c] = df_all[c].astype(str)
        df_all['unique_key'] = df_all.apply(lambda r: make_match_key(r[item], r[option]), axis=1)
        
        f_c1, f_c2, f_c3 = st.columns([2, 1, 1])
        search_q = f_c1.text_input("🔍 상품명 검색")
        filter_m = f_c2.selectbox("품절 필터", ["전체보기", "정상만", "품절만"], index=1)
        hist_date = f_c3.date_input("🗓️ 과거 입고확인 날짜", datetime.now())
        
        past_hist = load_history_from_gsheet()
        df_all['과거 리오더입고'] = 0
        if not past_hist.empty:
            try:
                past_hist['날짜'] = pd.to_datetime(past_hist['저장시간']).dt.date
                t_hist = past_hist[(past_hist['날짜'] == hist_date) & (past_hist['구분'] == "입고")].copy()
                if not t_hist.empty:
                    t_hist['k_tmp'] = t_hist.apply(lambda r: make_match_key(r['상품명'], r['옵션']), axis=1)
                    df_all['과거 리오더입고'] = df_all['unique_key'].map(t_hist.groupby('k_tmp')['수량'].sum().to_dict()).fillna(0).astype(int)
            except: pass

        v7, v3 = safe_num(df_all[t7day]), safe_num(df_all[t3day])
        df_all['일판매량'] = (v7 / 7 if v7.sum() > 0 else v3 / 3).round(1)
        df_all['권장발주량'] = ((df_all['일판매량'] * (lt + ss)) - (safe_num(df_all[avail]) + safe_num(df_all['리오더 수량']))).clip(lower=0).round(0).astype(int)
        df_all['리오더입고수량'] = 0
        df_work = df_all.copy()
        
        if filter_m == "정상만": df_work = df_work[~df_work[sold_out].astype(str).str.contains('품절', na=False)]
        elif filter_m == "품절만": df_work = df_work[df_work[sold_out].astype(str).str.contains('품절', na=False)]
        if search_q: df_work = df_work[df_work[item].astype(str).str.contains(search_q, case=False, na=False)]
        
        disp4 = [reg_date, sold_out, vendor, v_item, item, option, stock, avail, "리오더 수량", "리오더입고수량", "과거 리오더입고", "일판매량", "권장발주량", "unique_key"]
        edited4 = st.data_editor(df_work[disp4], width='stretch', hide_index=True, key="ed4", column_config={"unique_key": None})
        
        if st.button("💾 데이터 저장"):
            for _, row in edited4.iterrows():
                target_key = row["unique_key"]
                idx_list = st.session_state.df_raw[st.session_state.df_raw.apply(lambda r: make_match_key(r[item], r[option]), axis=1) == target_key].index
                if not idx_list.empty:
                    o_idx = idx_list[0]
                    st.session_state.df_raw.at[o_idx, "리오더 수량"] = int(row["리오더 수량"])
                    if int(row["리오더입고수량"]) > 0:
                        st.session_state.df_raw.at[o_idx, "리오더 수량"] = max(0, int(row["리오더 수량"]) - int(row["리오더입고수량"]))
                        save_history_to_gsheet(pd.DataFrame([[row[item], row[option], row["리오더입고수량"]]], columns=['상품명', '옵션', '수량']), "입고")
            save_df = st.session_state.df_raw[[item, option, '리오더 수량']].rename(columns={item:'상품명', option:'옵션'})
            if save_reorder_data(save_df): st.success("✅ 저장 완료!"); st.rerun()

        st.divider()
        df_final_base = df_work.copy()
        def get_stat(r):
            total_inv = safe_num(r[avail]) + safe_num(r['리오더 수량'])
            if r['일판매량'] <= 0: return "✅ 정상"
            days_left = total_inv / r['일판매량']
            return "🚨 긴급" if days_left < 3 else "⚠️ 주의" if days_left < 7 else "✅ 정상"
        
        df_final_base['상태'] = df_final_base.apply(get_stat, axis=1)
        df_final_base['추가 리오더'] = df_final_base['unique_key'].map(st.session_state.extra_order_dict).fillna(0).astype(int)
        df_final_base['최종발주량'] = df_final_base['권장발주량'] + df_final_base['추가 리오더']
        
        stat_filter = st.multiselect("상태 필터", ["🚨 긴급", "⚠️ 주의", "✅ 정상"], default=["🚨 긴급", "⚠️ 주의", "✅ 정상"])
        df_final = df_final_base[df_final_base['상태'].isin(stat_filter)].copy()
        df_final = df_final.sort_values(by=[item])
        
        disp5 = ["상태", reg_date, item, option, v_item, vendor, avail, "리오더 수량", "추가 리오더", "과거 리오더입고", "권장발주량", "최종발주량", "unique_key"]
        edited5 = st.data_editor(df_final[disp5], width='stretch', hide_index=True, key="ed5", column_config={"unique_key": None})
        for _, row in edited5.iterrows():
            st.session_state.extra_order_dict[row["unique_key"]] = int(row["추가 리오더"])
            
        c1, c2 = st.columns(2)
        if c1.button("📄 발주 기록 저장", width='stretch'):
            order_data = edited5[edited5['최종발주량'] > 0]
            if not order_data.empty:
                save_history_to_gsheet(order_data[[item, option, '최종발주량']].rename(columns={item:'상품명', option:'옵션', '최종발주량':'수량'}), "발주")
                st.success("발주 기록 저장 완료!")
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final[disp5[:-1]].to_excel(writer, index=False, sheet_name='발주리스트')
        c2.download_button(label="📥 엑셀 다운로드", data=output.getvalue(), file_name=f"Order_{datetime.now().strftime('%m%d_%H%M')}.xlsx", mime="application/vnd.ms-excel", width='stretch')

with tab2:
    st.subheader("📜 히스토리")
    st.dataframe(load_history_from_gsheet().sort_values(by="저장시간", ascending=False), width='stretch')
