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

# --- [4단계: 데이터 편집 및 재고 관리] ---
            st.divider()
            st.subheader("📊 4단계: 데이터 편집 및 재고 관리")
            
            df_work = st.session_state.df_raw.copy()

            # 1. 상단 UI 컨트롤러
            f_c1, f_c2, f_c3, f_c4 = st.columns([1.5, 1, 1, 1])
            search_q = f_c1.text_input("🔍 상품명 검색", key="search_v4_all")
            filter_m = f_c2.selectbox("품절 필터", ["전체보기", "정상만", "품절만"], index=1, key="filter_v4_all")
            show_all_items = f_c3.checkbox("QT/BE 상품 포함하기", value=True, key="show_all_v4_all")
            # [중요] 이 날짜의 입고 기록을 구글 시트에서 가져옵니다.
            hist_date_4 = f_c4.date_input("🗓️ 입고 기록 확인 날짜", datetime.now(), key="date_v4_all")

            # 2. 강력 매핑 함수 (공백, 대소문자, QT/BE 제거)
            def total_clean(n):
                import re
                if pd.isna(n): return ""
                text = str(n).strip().upper()
                text = re.sub(r'^(QT|BE)\s*', '', text)
                return text.replace(" ", "") # 모든 공백 제거

            # 원본 데이터에 고유 키 생성
            df_work['unique_key'] = df_work[item].apply(total_clean) + df_work[option].astype(str).str.replace(" ", "").str.upper()

            # 3. 구글 시트에서 '과거 리오더입고' 데이터 매핑
            past_hist = load_history_from_gsheet()
            df_work['과거 리오더입고'] = 0
            df_work['리오더입고수량'] = 0 # 오늘 입력용 칸
            
            if not past_hist.empty and '구분' in past_hist.columns:
                past_hist['날짜'] = past_hist['저장시간'].astype(str).str.split(' ').str[0]
                target_date_str = hist_date_4.strftime("%Y-%m-%d")
                
                # 선택한 날짜의 '입고' 데이터만 추출
                t_hist = past_hist[(past_hist['날짜'] == target_date_str) & (past_hist['구분'] == "입고")].copy()
                
                if not t_hist.empty:
                    # 시트 데이터도 똑같은 규칙으로 키 생성
                    t_hist['k_tmp'] = t_hist['상품명'].apply(total_clean) + t_hist['옵션'].astype(str).str.replace(" ", "").str.upper()
                    in_map = t_hist.groupby('k_tmp')['수량'].sum().to_dict()
                    # 매핑 실행 (여기서 숫자가 채워짐)
                    df_work['과거 리오더입고'] = df_work['unique_key'].map(in_map).fillna(0).astype(int)

            # 4. 수치 계산 로직
            def safe_num(val):
                res = pd.to_numeric(val, errors='coerce')
                return res.fillna(0) if hasattr(res, 'fillna') else (0 if pd.isna(res) else res)

            v7 = safe_num(df_work[t7day])
            v3 = safe_num(df_work[t3day])
            df_work['일판매량'] = (v7 / 7 if v7.sum() > 0 else v3 / 3).round(0).astype(int)
            df_work['권장발주량'] = ((df_work['일판매량'] * (lt + ss)) - (safe_num(df_work[avail]) + safe_num(df_work['리오더 수량']))).clip(lower=0).astype(int)

            # 5. 필터링 적용
            if filter_m == "정상만":
                df_work = df_work[~df_work[sold_out].astype(str).str.contains('품절', na=False)]
            elif filter_m == "품절만":
                df_work = df_work[df_work[sold_out].astype(str).str.contains('품절', na=False)]
            if not show_all_items:
                df_work = df_work[~df_work[item].astype(str).str.contains('QT|qt|BE|be', na=False)]
            if search_q:
                df_work = df_work[df_work[item].astype(str).str.contains(search_q, case=False, na=False)]

            # 6. 테이블 출력 (요청하신 순서)
            if not df_work.empty:
                display_cols = [sold_out, vendor, item, option, stock, avail, "리오더 수량", "리오더입고수량", "과거 리오더입고", t3day, "일판매량", "권장발주량"]
                valid_cols = [c for c in display_cols if c in df_work.columns or c in ["리오더입고수량", "과거 리오더입고"]]
                df_view = df_work[valid_cols].copy()
                for c in df_view.columns: df_view[c] = df_view[c].astype(str)
                
                def on_edit_4():
                    changes = st.session_state["editor_v4_all"]["edited_rows"]
                    for r_idx_str, change in changes.items():
                        idx = int(r_idx_str)
                        orig_idx = df_work.index[idx]
                        if "리오더 수량" in change:
                            st.session_state.df_raw.at[orig_idx, "리오더 수량"] = int(change["리오더 수량"])
                        if "리오더입고수량" in change:
                            in_qty = int(change["리오더입고수량"])
                            curr_re = st.session_state.df_raw.at[orig_idx, "리오더 수량"]
                            st.session_state.df_raw.at[orig_idx, "리오더 수량"] = max(0, curr_re - in_qty)
                            save_history_to_gsheet(pd.DataFrame([[df_work.at[orig_idx, item], df_work.at[orig_idx, option], in_qty]], columns=['상품명', '옵션', '수량']), log_type="입고")
                    save_reorder_data(st.session_state.df_raw[[item, option, '리오더 수량']].rename(columns={item:'상품명', option:'옵션'}))
                    st.rerun()

                st.data_editor(df_view, use_container_width=True, key="editor_v4_all", on_change=on_edit_4,
                               column_config={c: st.column_config.TextColumn(c) for c in df_view.columns}, hide_index=True)
                
# --- [5단계: 최종 발주 리스트 요약] ---
            st.divider()
            st.subheader("📋 5단계: 최종 발주 리스트 요약")
            
            # 1. 상단 컨트롤러 (필터 및 날짜)
            c5_1, c5_2 = st.columns([2, 1])
            s_filter = c5_1.selectbox("🎯 상태 필터", ["🚨긴급 + ⚠️주의 우선", "🚨 긴급만 보기", "✅ 정상 포함 전체보기"], index=0, key="s_filter_v5_final")
            # 4단계 날짜와 연동되지만 여기서도 변경 가능
            hist_date_5 = c5_2.date_input("🗓️ 입고 기록 확인 날짜 (연동)", value=hist_date_4, key="date_5_v5_final")

            # 2. 데이터 준비 및 강력 키 생성
            to_order = df_work.copy()
            
            def total_clean(n):
                import re
                if pd.isna(n): return ""
                text = str(n).strip().upper()
                text = re.sub(r'^(QT|BE)\s*', '', text)
                return text.replace(" ", "")

            to_order['unique_key'] = to_order[item].apply(total_clean) + to_order[option].astype(str).str.replace(" ", "").str.upper()

            # 3. 과거 입고 데이터 매핑 (KeyError 방지 로직)
            to_order['과거 리오더입고'] = 0
            past_hist = load_history_from_gsheet()
            if not past_hist.empty:
                if '저장시간' in past_hist.columns:
                    past_hist['날짜'] = past_hist['저장시간'].astype(str).str.split(' ').str[0]
                    target_date_5 = hist_date_5.strftime("%Y-%m-%d")
                    if '구분' in past_hist.columns:
                        t_hist_5 = past_hist[(past_hist['날짜'] == target_date_5) & (past_hist['구분'] == "입고")].copy()
                        if not t_hist_5.empty:
                            t_hist_5['k_tmp'] = t_hist_5['상품명'].apply(total_clean) + t_hist_5['옵션'].astype(str).str.replace(" ", "").str.upper()
                            in_map_5 = t_hist_5.groupby('k_tmp')['수량'].sum().to_dict()
                            to_order['과거 리오더입고'] = to_order['unique_key'].map(in_map_5).fillna(0).astype(int)

            # 4. 수치 계산 (추가발주 및 최종발주량)
            if "extra_order_dict" not in st.session_state:
                st.session_state.extra_order_dict = {}
            
            to_order['추가발주수량'] = to_order['unique_key'].map(st.session_state.extra_order_dict).fillna(0).astype(int)
            to_order['최종발주량'] = to_order['권장발주량'].astype(int) + to_order['추가발주수량'].astype(int)

            # 5. 상태 판별 및 정렬
            def get_final_status(r):
                def safe_n(v):
                    res = pd.to_numeric(v, errors='coerce')
                    return res if not pd.isna(res) else 0
                total = safe_n(r[avail]) + safe_n(r['리오더 수량'])
                daily = r['일판매량']
                if daily > 0:
                    if total < (daily * 3): return "🚨 긴급"
                    if total < (daily * 5): return "⚠️ 주의"
                return "✅ 정상"
            
            to_order['상태'] = to_order.apply(get_final_status, axis=1)
            status_rank = {"🚨 긴급": 0, "⚠️ 주의": 1, "✅ 정상": 2}
            to_order['rank'] = to_order['상태'].map(status_rank)
            to_order = to_order.sort_values(by='rank').drop(columns=['rank'])

            # 6. 필터링 적용
            df_final = to_order.copy()
            if s_filter == "🚨긴급 + ⚠️주의 우선":
                df_final = df_final[df_final['상태'].isin(["🚨 긴급", "⚠️ 주의"]) | (df_final['최종발주량'] > 0)]
            elif s_filter == "🚨 긴급만 보기":
                df_final = df_final[df_final['상태'] == "🚨 긴급"]

            # 7. 테이블 출력 (항목 순서 보장)
            if not df_final.empty:
                disp_final = ["상태", item, option, vendor, avail, "리오더 수량", "추가발주수량", "과거 리오더입고", "권장발주량", "최종발주량"]
                valid_f_cols = [c for c in disp_final if c in df_final.columns or c in ["과거 리오더입고", "추가발주수량"]]
                df_final_view = df_final[valid_f_cols].copy()
                for c in df_final_view.columns: df_final_view[c] = df_final_view[c].astype(str)
                
                def on_edit_5():
                    edits = st.session_state["editor_v5_final"]["edited_rows"]
                    for r_idx_str, change in edits.items():
                        if "추가발주수량" in change:
                            r_key = df_final.iloc[int(r_idx_str)]['unique_key']
                            st.session_state.extra_order_dict[r_key] = int(change["추가발주수량"])
                    st.rerun()

                st.data_editor(df_final_view, use_container_width=True, key="editor_v5_final", on_change=on_edit_5,
                               column_config={c: st.column_config.TextColumn(c) for c in df_final_view.columns}, hide_index=True)

                # --- [핵심: 버튼 섹션] ---
                st.write("") 
                btn_c1, btn_c2 = st.columns(2)
                
                # 구글 시트 저장 버튼
                if btn_c1.button("💾 구글 시트에 최종 발주 기록 저장", use_container_width=True):
                    to_save = df_final[df_final['최종발주량'].astype(int) > 0].copy()
                    if not to_save.empty:
                        save_df = to_save[[item, option, '최종발주량']].rename(columns={item: '상품명', option: '옵션', '최종발주량': '수량'})
                        if save_history_to_gsheet(save_df, log_type="발주"):
                            st.success(f"✅ {len(save_df)}건 저장 완료!")
                        else:
                            st.error("❌ 시트 저장 실패")
                    else:
                        st.warning("발주 수량이 있는 상품이 없습니다.")

                # 엑셀 다운로드 버튼
                csv_data = df_final.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
                btn_c2.download_button(
                    label="📥 엑셀(CSV) 다운로드",
                    data=csv_data,
                    file_name=f"최종발주서_{datetime.now().strftime('%m%d_%H%M')}.csv",
                    mime='text/csv',
                    use_container_width=True
                )
            else:
                st.info("💡 표시할 발주 데이터가 없습니다.")
                
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
