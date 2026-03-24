import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
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
        
        # 사장님 요청: 매핑 설정 양쪽 5개씩 정렬
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
        if st.button("📊 분석 실행", width='stretch'):
            st.session_state.analyzed = True
            st.rerun()
            
if st.session_state.get('analyzed'):
        # --- [4단계: 데이터 편집 및 재고 관리 - 클린 버전] ---
        st.divider()
        st.subheader("📊 4단계: 데이터 편집 및 재고 관리")
        
        df_work = st.session_state.df_raw.copy()

        # 1. 상단 UI 컨트롤러
        f_c1, f_c2, f_c3 = st.columns([2, 1, 1])
        search_q = f_c1.text_input("🔍 상품명 검색", key="search_v4_clean")
        filter_m = f_c2.selectbox("품절 필터", ["전체보기", "정상만", "품절만"], index=1, key="filter_v4_clean")
        hist_date_4 = f_c3.date_input("🗓️ 입고 기록 확인 날짜", datetime.now(), key="date_v4_clean")

        # 2. 매핑 키 생성 함수
        def simple_key(n):
            return str(n).strip().replace(" ", "").upper() if not pd.isna(n) else ""

        df_work['unique_key'] = df_work[item].apply(simple_key) + df_work[option].apply(simple_key)

        # 3. 구글 시트 데이터 매핑
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

        # 4. 수치 계산
        v7 = safe_num(df_work[t7day])
        v3 = safe_num(df_work[t3day])
        df_work['일판매량'] = (v7 / 7 if v7.sum() > 0 else v3 / 3).round(0).astype(int)
        df_work['권장발주량'] = ((df_work['일판매량'] * (lt + ss)) - (safe_num(df_work[avail]) + safe_num(df_work['리오더 수량']))).clip(lower=0).astype(int)

        # 5. 필터링
        if filter_m == "정상만":
            df_work = df_work[~df_work[sold_out].astype(str).str.contains('품절', na=False)]
        elif filter_m == "품절만":
            df_work = df_work[df_work[sold_out].astype(str).str.contains('품절', na=False)]
        
        if search_q:
            df_work = df_work[df_work[item].astype(str).str.contains(search_q, case=False, na=False)]

        # 6. 테이블 출력 및 자동 저장 로직
        if not df_work.empty:
            display_cols = [sold_out, vendor, item, option, stock, avail, "리오더 수량", "리오더입고수량", "과거 리오더입고", t3day, "일판매량", "권장발주량"]
            valid_cols = [c for c in display_cols if c in df_work.columns or c in ["리오더입고수량", "과거 리오더입고"]]
            df_view = df_work[valid_cols].copy()
            for c in df_view.columns: df_view[c] = df_view[c].astype(str)
            
            def on_edit_4():
                changes = st.session_state["editor_v4_clean"]["edited_rows"]
                for r_idx_str, change in changes.items():
                    idx = int(r_idx_str)
                    orig_idx = df_work.index[idx]
                    if "리오더 수량" in change:
                        st.session_state.df_raw.at[orig_idx, "리오더 수량"] = int(change["리오더 수량"])
                    if "리오더입고수량" in change:
                        in_qty = int(change["리오더입고수량"])
                        st.session_state.df_raw.at[orig_idx, "리오더 수량"] = max(0, int(st.session_state.df_raw.at[orig_idx, "리오더 수량"]) - in_qty)
                        save_history_to_gsheet(pd.DataFrame([[df_work.at[orig_idx, item], df_work.at[orig_idx, option], in_qty]], columns=['상품명', '옵션', '수량']), log_type="입고")
                
                save_df = st.session_state.df_raw[[item, option, '리오더 수량']].rename(columns={item:'상품명', option:'옵션'})
                save_reorder_data(save_df)
                st.rerun()

            st.data_editor(df_view, use_container_width=True, key="editor_v4_clean", on_change=on_edit_4, hide_index=True)

        # --- [5단계: 최종 발주 리스트 요약] ---
        st.divider()
        st.subheader("📋 5단계: 최종 발주 리스트 요약")
        
        # 1. 상단 컨트롤러
        c5_1, c5_2 = st.columns([2, 1])
        s_filter = c5_1.selectbox("🎯 상태 필터", ["🚨긴급 + ⚠️주의 우선", "🚨 긴급만 보기", "✅ 정상 포함 전체보기"], index=0, key="s_filter_v5_final_v3")
        hist_date_5 = c5_2.date_input("🗓️ 입고 기록 확인 날짜 (연동)", value=hist_date_4, key="date_5_v5_final_v3")

        # 2. 데이터 준비
        to_order = df_work.copy()
        to_order['unique_key'] = to_order[item].apply(simple_key) + to_order[option].apply(simple_key)

        # 3. 과거 입고 데이터 매핑 (방어 로직 포함)
        to_order['과거 리오더입고'] = 0
        if not past_hist.empty:
            try:
                if '저장시간' in past_hist.columns:
                    past_hist['날짜'] = past_hist['저장시간'].astype(str).str.split(' ').str[0]
                    target_date_5 = hist_date_5.strftime("%Y-%m-%d")
                    if '구분' in past_hist.columns:
                        t_hist_5 = past_hist[(past_hist['날짜'] == target_date_5) & (past_hist['구분'] == "입고")].copy()
                        if not t_hist_5.empty:
                            t_hist_5['k_tmp'] = t_hist_5['상품명'].apply(simple_key) + t_hist_5['옵션'].apply(simple_key)
                            in_map_5 = t_hist_5.groupby('k_tmp')['수량'].sum().to_dict()
                            to_order['과거 리오더입고'] = to_order['unique_key'].map(in_map_5).fillna(0).astype(int)
            except:
                pass

        # 4. 수치 계산
        if "extra_order_dict" not in st.session_state:
            st.session_state.extra_order_dict = {}
        
        to_order['추가발주수량'] = to_order['unique_key'].map(st.session_state.extra_order_dict).fillna(0).astype(int)
        to_order['최종발주량'] = to_order['권장발주량'].astype(int) + to_order['추가발주수량'].astype(int)

        # 5. 상태 판별 및 정렬
        def get_final_status(r):
            total = safe_num(r[avail]) + safe_num(r['리오더 수량'])
            daily = r['일판매량']
            if daily > 0:
                if total < (daily * 3): return "🚨 긴급"
                if total < (daily * 5): return "⚠️ 주의"
            return "✅ 정상"
        
        to_order['상태'] = to_order.apply(get_final_status, axis=1)
        status_rank = {"🚨 긴급": 0, "⚠️ 주의": 1, "✅ 정상": 2}
        to_order['rank'] = to_order['상태'].map(status_rank)
        to_order = to_order.sort_values(by='rank').drop(columns=['rank'])

        # 6. 필터링 및 출력
        df_final = to_order.copy()
        if s_filter == "🚨긴급 + ⚠️주의 우선":
            df_final = df_final[df_final['상태'].isin(["🚨 긴급", "⚠️ 주의"]) | (df_final['최종발주량'] > 0)]
        elif s_filter == "🚨 긴급만 보기":
            df_final = df_final[df_final['상태'] == "🚨 긴급"]

        if not df_final.empty:
            disp_final = ["상태", item, option, vendor, avail, "리오더 수량", "추가발주수량", "과거 리오더입고", "권장발주량", "최종발주량"]
            df_final_view = df_final[[c for c in disp_final if c in df_final.columns]].copy()
            for c in df_final_view.columns: df_final_view[c] = df_final_view[c].astype(str)
            
            def on_edit_5():
                edits = st.session_state["editor_v5_final_v3"]["edited_rows"]
                for r_idx_str, change in edits.items():
                    if "추가발주수량" in change:
                        r_key = df_final.iloc[int(r_idx_str)]['unique_key']
                        st.session_state.extra_order_dict[r_key] = int(change["추가발주수량"])
                st.rerun()

           # 5단계 데이터 에디터 출력
            st.data_editor(df_final_view, use_container_width=True, key="editor_v5_final_v3", on_change=on_edit_5, hide_index=True)

            # 7. 저장 및 다운로드 버튼
            st.write("")
            b1, b2 = st.columns(2)
            if b1.button("💾 구글 시트에 최종 발주 기록 저장", use_container_width=True, key="btn_save_v5"):
                to_save = df_final[df_final['최종발주량'].astype(int) > 0].copy()
                if not to_save.empty:
                    save_df = to_save[[item, option, '최종발주량']].rename(columns={item: '상품명', option: '옵션', '최종발주량': '수량'})
                    if save_history_to_gsheet(save_df, log_type="발주"):
                        st.success("✅ 저장 성공!")
                    else:
                        st.error("❌ 저장 실패")
                else:
                    st.warning("발주 수량이 있는 항목이 없습니다.")

            csv_data = df_final.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            b2.download_button(label="📥 엑셀(CSV) 다운로드", data=csv_data, file_name=f"발주서_{datetime.now().strftime('%m%d')}.csv", use_container_width=True, key="btn_dl_v5")

        else:
            # 5단계에서 표시할 데이터가 없을 때의 else (에러 났던 부분)
            st.info("💡 표시할 발주 데이터가 없습니다.")

        # --- [6단계: 과거 데이터 통합 조회] ---
        st.divider()
        st.subheader("📜 6단계: 과거 데이터 통합 조회")
        
        # 1. 상단 컨트롤러 (날짜 범위 선택)
        c6_1, c6_2 = st.columns(2)
        with c6_1:
            start_d = st.date_input("조회 시작 날짜", pd.Timestamp.now() - pd.Timedelta(days=7), key="s_date_v6_fix")
        with c6_2:
            end_d = st.date_input("조회 종료 날짜", pd.Timestamp.now(), key="e_date_v6_fix")

        # 2. 데이터 불러오기
        hist_all = load_history_from_gsheet()
        
        if not hist_all.empty:
            try:
                if '저장시간' in hist_all.columns:
                    hist_all['날짜_dt'] = pd.to_datetime(hist_all['저장시간'], errors='coerce').dt.date
                    df_filtered = hist_all[(hist_all['날짜_dt'] >= start_d) & (hist_all['날짜_dt'] <= end_d)].copy()
                    
                    if not df_filtered.empty:
                        tab_in, tab_out = st.tabs(["📥 입고 내역 기록", "📤 발주 내역 기록"])
                        
                        with tab_in:
                            if '구분' in df_filtered.columns:
                                in_data = df_filtered[df_filtered['구분'] == "입고"]
                                if not in_data.empty:
                                    st.dataframe(in_data[['저장시간', '상품명', '옵션', '수량']], use_container_width=True, hide_index=True)
                                    sum_in = in_data.groupby(['상품명', '옵션'])['수량'].sum().reset_index()
                                    st.write("📋 **해당 기간 상품별 입고 총합**")
                                    st.table(sum_in)
                                else:
                                    st.info("기간 내 입고 기록이 없습니다.")
                        
                        with tab_out:
                            if '구분' in df_filtered.columns:
                                out_data = df_filtered[df_filtered['구분'] == "발주"]
                                if not out_data.empty:
                                    st.dataframe(out_data[['저장시간', '상품명', '옵션', '수량']], use_container_width=True, hide_index=True)
                                    sum_out = out_data.groupby(['상품명', '옵션'])['수량'].sum().reset_index()
                                    st.write("📋 **해당 기간 상품별 발주 총합**")
                                    st.table(sum_out)
                                else:
                                    st.info("기간 내 발주 기록이 없습니다.")
                    else:
                        st.warning("선택하신 기간에 해당하는 데이터가 시트에 없습니다.")
            except Exception as e:
                st.error(f"데이터 조회 중 오류가 발생했습니다: {e}")
        else:
            st.info("💡 구글 시트에 저장된 기록이 없습니다.")

        # 3. 전체 데이터 다운로드
        if not hist_all.empty:
            st.write("---")
            csv_all = hist_all.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button(
                label="📥 전체 히스토리 다운로드 (CSV)",
                data=csv_all,
                file_name=f"전체기록_{datetime.now().strftime('%m%d')}.csv",
                use_container_width=True,
                key="btn_full_dl"
            )

# --- [🌙 탭 2: 동대문 사입 관리] ---
with tab2:
    st.subheader("🌙 동대문 사입 및 미납 관리")

    dong_file = st.file_uploader("동대문 주문 리스트 업로드", type=['xlsx', 'xls', 'csv'], key="dong_tab_upload")

    if dong_file:
        # 파일이 새로 올라왔을 때만 데이터 처리
        if "last_file_name" not in st.session_state or st.session_state.last_file_name != dong_file.name:
            # 엑셀/CSV 구분해서 읽기
            if dong_file.name.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(dong_file)
            else:
                df = pd.read_csv(dong_file)
            
            df.columns = df.columns.str.strip()
            
            # 필수 컬럼 체크 및 생성
            required_cols = ['선택', '품절', '상품명', '공급처', '공급처상품명', '정상재고', '가용재고', '판매수량', '발주수량', '가중율', '3일판매']
            for col in required_cols:
                if col not in df.columns:
                    if col in ['선택', '품절', '상품명', '공급처', '공급처상품명']:
                        df[col] = ""
                    else:
                        df[col] = 0
            
            # 숫자형 변환
            for col in ['정상재고', '가용재고', '3일판매']:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
            # 동대문 전용 발주 로직 적용
            df['판매수량'] = (df['정상재고'] - df['가용재고']).clip(lower=0)
            # 판매량에 따른 가중율 (10개 이상 2배, 6개 이상 1.5배 등)
            df['가중율'] = df['판매수량'].apply(lambda n: 2.0 if n >= 10 else (1.5 if n >= 6 else (1.2 if n >= 3 else 1.0)))
            df['발주수량'] = (df['판매수량'] * df['가중율']).astype(int)
            
            st.session_state.df_dong_current = df[required_cols]
            st.session_state.last_file_name = dong_file.name

        # 화면 출력 부분
        if "df_dong_current" in st.session_state:
            df_display = st.session_state.df_dong_current.copy()
            
            # 검색 기능
            search_query = st.text_input("🔍 상품명 검색 (사입)")
            if search_query:
                df_display = df_display[df_display['상품명'].astype(str).str.contains(search_query, case=False, na=False)]
            
            # 데이터 에디터 (선택 및 수량 수정 가능)
            df_display['선택'] = df_display['선택'].apply(lambda x: True if x is True or x == "True" else False)
            
            edited_df = st.data_editor(
                df_display, 
                use_container_width=True, 
                key="dong_editor",
                column_config={
                    "선택": st.column_config.CheckboxColumn("선택", default=False),
                    "발주수량": st.column_config.NumberColumn("발주수량", min_value=0)
                },
                hide_index=True
            )
            
            st.divider()
            
            # 하단 컨트롤러
            c1, c2, c3 = st.columns([1, 1, 1])
            add_val = c1.number_input("➕ 추가 수량", value=1, min_value=1, key="dong_add_val")
            
            if c2.button("🚀 선택 상품 수량 더하기", use_container_width=True):
                # 에디터에서 선택된 인덱스 찾기
                # 실제 세션 데이터에 반영
                for i, row in edited_df.iterrows():
                    if row['선택']:
                        # 원본 데이터의 인덱스를 찾아 업데이트
                        st.session_state.df_dong_current.at[i, '발주수량'] += add_val
                st.success(f"선택 항목에 {add_val}개씩 추가되었습니다.")
                st.rerun()
            
            # 다운로드 버튼
            csv_dong = edited_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            c3.download_button(
                label="📥 사입 리스트 다운로드 (CSV)", 
                data=csv_dong, 
                file_name=f"동대문사입_{datetime.now().strftime('%m%d')}.csv", 
                mime="text/csv",
                use_container_width=True
            )
    else:
        st.info("💡 동대문 발주용 엑셀 파일을 업로드해주세요.")
