import streamlit as st
import pandas as pd
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
    except:
        return None

def save_reorder_data(df):
    try:
        sheet = get_sheet().sheet1
        sheet.clear()
        sheet.update([df.columns.values.tolist()] + df.fillna(0).values.tolist())
    except:
        pass

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

# --- [앱 최상단 혹은 세션 초기화 구역에 추가] ---
if "extra_order_dict" not in st.session_state:
    st.session_state.extra_order_dict = {}

# --- [2. 앱 설정 및 세션 초기화] ---
st.set_page_config(layout="wide", page_title="재고 관리 시스템")
st.title("📦 통합 재고 관리 시스템")

if 'analyzed' not in st.session_state: st.session_state.analyzed = False

tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 관리"])

with tab1:
    st.subheader("📁 데이터 업로드 (제작상품)")
    if st.button("🔄 제작상품 데이터 초기화"):
        # 모든 세션을 지우지 않고 데이터 관련 키만 선별 삭제 (설정 유지)
        for key in ['df_raw', 'analyzed', 'last_filename']:
            if key in st.session_state: del st.session_state[key]
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
                    df_new['k_tmp'] = df_new[t_item].astype(str).str.strip() + df_new[t_opt].astype(str).str.strip()
                    gs_data['k_tmp'] = gs_data['상품명'].astype(str).str.strip() + gs_data['옵션'].astype(str).str.strip()
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

        # --- 1단계: 매핑 설정 ---
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
        t7day = c2.selectbox("7일 발주합계", cols, index=find_idx(cols, ['7일', '1주']))

        # --- 2~3단계: 분석 설정 ---
        st.subheader("🚀 2~3단계: 분석 설정")
        l1, l2 = st.columns(2)
        lt = l1.number_input("리드타임 (일)", value=10)
        ss = l2.number_input("안전재고 (일 수)", value=7)
        if st.button("📊 분석 실행", use_container_width=True):
            st.session_state.analyzed = True
            st.rerun()

if st.session_state.analyzed:
            # --- [필수 함수 정의: AttributeError 방지형] ---
            def safe_to_num(val):
                """값 하나 혹은 시리즈를 숫자로 안전하게 변환"""
                import pandas as pd
                res = pd.to_numeric(val, errors='coerce')
                # 결과가 판다스 객체면 fillna 사용, 아니면 일반 숫자 처리
                if hasattr(res, 'fillna'):
                    return res.fillna(0)
                return 0 if pd.isna(res) else res

if st.session_state.analyzed:
            # --- [필수 함수 정의: AttributeError 및 유실 방지] ---
            def safe_num(val):
                import pandas as pd
                res = pd.to_numeric(val, errors='coerce')
                return res.fillna(0) if hasattr(res, 'fillna') else (0 if pd.isna(res) else res)

if st.session_state.analyzed:
            # [안전장치] 추가발주수량 딕셔너리가 없으면 생성
            if "extra_order_dict" not in st.session_state:
                st.session_state.extra_order_dict = {}

if st.session_state.analyzed:
            # --- [필수 함수 정의] ---
            def safe_num(val):
                res = pd.to_numeric(val, errors='coerce')
                return res.fillna(0) if hasattr(res, 'fillna') else (0 if pd.isna(res) else res)

KeyError: This app has encountered an error. The original error message is redacted to prevent data leaks. Full error details have been recorded in the logs (if you're on Streamlit Cloud, click on 'Manage app' in the lower right of your app).
Traceback:
File "/mount/src/order-system/app.py", line 251, in <module>
    to_order['추가발주수량'] = to_order['unique_key'].map(st.session_state.extra_order_dict).fillna(0).astype(int)
                               ~~~~~~~~^^^^^^^^^^^^^^
File "/home/adminuser/venv/lib/python3.14/site-packages/pandas/core/frame.py", line 4113, in __getitem__
    indexer = self.columns.get_loc(key)
File "/home/adminuser/venv/lib/python3.14/site-packages/pandas/core/indexes/base.py", line 3819, in get_loc
    raise KeyError(key) from err


# --- [4단계: 데이터 편집 및 재고 관리 - 최종 복구본] ---
            st.divider()
            st.subheader("📊 4단계: 데이터 편집 및 재고 관리")
            
            df_work = st.session_state.df_raw.copy()

            # 1. 상단 UI 컨트롤러
            f_c1, f_c2, f_c3, f_c4 = st.columns([1.5, 1, 1, 1])
            search_q = f_c1.text_input("🔍 상품명 검색", key="search_v4_final")
            filter_m = f_c2.selectbox("품절 필터", ["전체보기", "정상만", "품절만"], index=1, key="filter_v4_final")
            show_all_items = f_c3.checkbox("QT/BE 상품 포함하기", value=True, key="show_all_v4_final")
            hist_date_4 = f_c4.date_input("🗓️ 입고 날짜", datetime.now(), key="date_v4_final")

            # 2. 클린 상품명 및 고유 키 생성
            def clean_name(n):
                import re
                return re.sub(r'^(QT|BE|qt|be)\s*', '', str(n)).strip()

            df_work['clean_item'] = df_work[item].apply(clean_name)
            df_work['unique_key'] = df_work['clean_item'] + df_work[option].astype(str).str.strip()

            # 3. 수치 계산을 위한 안전한 숫자 변환 함수
            def safe_num(val):
                res = pd.to_numeric(val, errors='coerce')
                return res.fillna(0) if hasattr(res, 'fillna') else (0 if pd.isna(res) else res)

            # 일판매량 및 권장발주량 계산
            v7 = safe_num(df_work[t7day])
            v3 = safe_num(df_work[t3day])
            df_work['일판매량'] = (v7 / 7 if v7.sum() > 0 else v3 / 3).round(0).astype(int)
            df_work['권장발주량'] = ((df_work['일판매량'] * (lt + ss)) - (safe_num(df_work[avail]) + safe_num(df_work['리오더 수량']))).clip(lower=0).astype(int)

            # 4. 필터링 로직 적용
            # (1) 품절 필터
            if sold_out in df_work.columns:
                if filter_m == "정상만":
                    df_work = df_work[~df_work[sold_out].astype(str).str.contains('품절', na=False)]
                elif filter_m == "품절만":
                    df_work = df_work[df_work[sold_out].astype(str).str.contains('품절', na=False)]

            # (2) QT/BE 제외 필터
            if not show_all_items:
                df_work = df_work[~df_work[item].astype(str).str.contains('QT|qt|BE|be', na=False)]

            # (3) 검색어 필터
            if search_q:
                df_work = df_work[df_work[item].astype(str).str.contains(search_q, case=False, na=False)]

            # 5. 테이블 출력
            if not df_work.empty:
                # 항목 복구: 품절상태, 제조사, 상품명, 옵션, 자체품번, 가용재고, 리오더수량, 입고수량, 3일판매량, 일판매량, 권장발주량
                display_list = [sold_out, vendor, item, option, vendor_item, avail, "리오더 수량", "리오더입고수량", t3day, "일판매량", "권장발주량"]
                valid_cols = [c for c in display_list if c in df_work.columns or c in ["리오더입고수량"]]
                
                # 리오더입고수량이 컬럼에 없으면 0으로 생성
                if "리오더입고수량" not in df_work.columns:
                    df_work["리오더입고수량"] = 0

                df_view = df_work[valid_cols].copy()
                for c in df_view.columns:
                    df_view[c] = df_view[c].astype(str)
                
                def on_edit_4():
                    changes = st.session_state["editor_v4_final"]["edited_rows"]
                    for r_idx_str, change in changes.items():
                        idx = int(r_idx_str)
                        orig_idx = df_work.index[idx]
                        if "리오더 수량" in change:
                            st.session_state.df_raw.at[orig_idx, "리오더 수량"] = int(change["리오더 수량"])
                    save_reorder_data(st.session_state.df_raw[[item, option, '리오더 수량']].rename(columns={item:'상품명', option:'옵션'}))
                    st.rerun()

                st.data_editor(
                    df_view, 
                    use_container_width=True, 
                    key="editor_v4_final", 
                    on_change=on_edit_4,
                    column_config={c: st.column_config.TextColumn(c) for c in df_view.columns},
                    hide_index=True
                )
            else:
                st.info("💡 표시할 데이터가 없습니다. 필터를 조정해 보세요.")
                
# --- [5단계: 최종 발주 리스트 요약] ---
            st.divider()
            st.subheader("📋 5단계: 최종 발주 리스트 요약")
            
            c5_1, c5_2 = st.columns([2, 1])
            # 초기 로딩 시 긴급/주의가 먼저 보이도록 설정
            s_filter = c5_1.selectbox("🎯 상태 필터", ["🚨긴급 + ⚠️주의 우선", "🚨 긴급만 보기", "✅ 정상 포함 전체보기"], index=0, key="s_filter_v5")
            hist_date_5 = c5_2.date_input("🗓️ 입고 기록 확인 날짜 (연동)", value=hist_date_4, key="date_5_v5")

            # 0. 데이터 복사 및 Key 재정의 (KeyError 방지 핵심)
            to_order = df_work.copy()
            
            # [수정] unique_key가 없는 경우를 대비해 여기서 다시 생성합니다.
            to_order['unique_key'] = to_order[item].astype(str).str.strip() + to_order[option].astype(str).str.strip()

            # 1. 수치 계산 및 상태 판별
            def safe_num(val):
                res = pd.to_numeric(val, errors='coerce')
                return res.fillna(0) if hasattr(res, 'fillna') else (0 if pd.isna(res) else res)

            # 추가발주수량 매핑 (st.session_state.extra_order_dict 사용)
            if "extra_order_dict" not in st.session_state:
                st.session_state.extra_order_dict = {}
                
            to_order['추가발주수량'] = to_order['unique_key'].map(st.session_state.extra_order_dict).fillna(0).astype(int)
            to_order['최종발주량'] = to_order['권장발주량'] + to_order['추가발주수량']

            def get_final_status(r):
                total = safe_num(r[avail]) + safe_num(r['리오더 수량'])
                daily = r['일판매량']
                if daily > 0:
                    if total < (daily * 3): return "🚨 긴급"
                    if total < (daily * 5): return "⚠️ 주의"
                return "✅ 정상"
            
            to_order['상태'] = to_order.apply(get_final_status, axis=1)

            # 2. 우선순위 정렬 (긴급 -> 주의 -> 정상)
            status_rank = {"🚨 긴급": 0, "⚠️ 주의": 1, "✅ 정상": 2}
            to_order['rank'] = to_order['상태'].map(status_rank)
            to_order = to_order.sort_values(by='rank').drop(columns=['rank'])

            # 3. 필터링 로직
            if s_filter == "🚨긴급 + ⚠️주의 우선":
                df_final = to_order[to_order['상태'].isin(["🚨 긴급", "⚠️ 주의"]) | (to_order['최종발주량'] > 0)].copy()
            elif s_filter == "🚨 긴급만 보기":
                df_final = to_order[to_order['상태'] == "🚨 긴급"].copy()
            else:
                df_final = to_order.copy()

            # 4. 결과 출력
            if not df_final.empty:
                disp_final = ["상태", item, option, vendor, avail, "리오더 수량", "추가발주수량", "권장발주량", "최종발주량"]
                # 존재하는 컬럼만 선택
                valid_final_cols = [c for c in disp_final if c in df_final.columns]
                df_final_view = df_final[valid_final_cols].copy()
                
                # 왼쪽 정렬용 문자열 변환
                for c in df_final_view.columns:
                    df_final_view[c] = df_final_view[c].astype(str)
                
                def on_edit_5():
                    edits = st.session_state["final_editor_v5"]["edited_rows"]
                    for r_idx_str, change in edits.items():
                        if "추가발주수량" in change:
                            # 정렬된 상태이므로 iloc로 정확한 key 추출
                            r_key = df_final.iloc[int(r_idx_str)]['unique_key']
                            try:
                                st.session_state.extra_order_dict[r_key] = int(change["추가발주수량"])
                            except: pass
                    st.rerun()

                st.data_editor(
                    df_final_view, 
                    use_container_width=True, 
                    key="final_editor_v5", 
                    on_change=on_edit_5,
                    column_config={c: st.column_config.TextColumn(c) for c in df_final_view.columns},
                    hide_index=True
                )
                
                # 저장 및 다운로드 버튼
                c_b1, c_b2 = st.columns(2)
                if c_b1.button("💾 구글 시트에 최종 발주 기록 저장", use_container_width=True):
                    save_df = df_final[df_final['최종발주량'].astype(int) > 0][[item, option, '최종발주량']]
                    if save_history_to_gsheet(save_df):
                        st.success("✅ 저장 완료!")
                
                csv = df_final.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
                c_b2.download_button("📥 엑셀 다운로드", csv, f"최종발주서_{datetime.now().strftime('%m%d')}.csv", use_container_width=True)
            else:
                st.info("💡 현재 발주할 상품이 없습니다.")
        # 6단계: 과거 확인

            st.divider()

            st.subheader("📜 6단계: 과거 데이터 확인")

            if st.button("🔄 기록 불러오기"): st.session_state.db_history = load_history_from_gsheet()

            if 'db_history' in st.session_state and not st.session_state.db_history.empty:

                df_hist = st.session_state.db_history

                df_hist['날짜'] = df_hist['저장시간'].astype(str).str.split(' ').str[0]

                sel_date = st.date_input("날짜 선택", datetime.now())

                day_data = df_hist[df_hist['날짜'] == sel_date.strftime("%Y-%m-%d")]

                if not day_data.empty:

                    sel_time = st.selectbox("⏰ 시간 선택", sorted(day_data['저장시간'].unique(), reverse=True))

                    st.dataframe(day_data[day_data['저장시간'] == sel_time].drop(columns=['날짜']), use_container_width=True)



# --- [🌙 탭 2: 동대문 사입 관리] ---

with tab2:

    st.subheader("🌙 동대문 사입 및 미납 관리")

    dong_file = st.file_uploader("동대문 주문 리스트 업로드", type=['xlsx', 'csv'], key="dong_tab_upload")

    if dong_file:

        if "last_file_name" not in st.session_state or st.session_state.last_file_name != dong_file.name:

            df = pd.read_excel(dong_file)

            df.columns = df.columns.str.strip()

            required_cols = ['선택', '품절', '상품명', '공급처', '공급처상품명', '정상재고', '가용재고', '판매수량', '발주수량', '가중율', '3일판매']

            for col in required_cols:

                if col not in df.columns: df[col] = 0 if col not in ['선택', '품절', '상품명', '공급처', '공급처상품명'] else ""

            for col in ['정상재고', '가용재고', '3일판매']: df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            df['판매수량'] = (df['정상재고'] - df['가용재고']).clip(lower=0)

            df['가중율'] = df['판매수량'].apply(lambda n: 2.0 if n >= 10 else (1.5 if n >= 6 else (1.2 if n >= 3 else 1.0)))

            df['발주수량'] = (df['판매수량'] * df['가중율']).astype(int)

            st.session_state.df_dong_current = df[required_cols]

            st.session_state.last_file_name = dong_file.name



        df_display = st.session_state.df_dong_current.copy()

        search_query = st.text_input("상품명 검색 (사입)")

        if search_query: df_display = df_display[df_display['상품명'].astype(str).str.contains(search_query, case=False, na=False)]

        

        df_display['선택'] = df_display['선택'].astype(bool)

        edited_df = st.data_editor(df_display, use_container_width=True, key="dong_editor")

        

        st.divider()

        c1, c2, c3 = st.columns(3)

        add_val = c1.number_input("추가 수량", value=1, min_value=1)

        if c2.button("🚀 선택 상품 수량 더하기"):

            selected = edited_df[edited_df['선택'] == True].index

            for idx in selected: st.session_state.df_dong_current.at[idx, '발주수량'] += add_val

            st.rerun()

        csv = edited_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')

        c3.download_button("📥 엑셀 다운로드", csv, "사입리스트.csv", "text/csv")
