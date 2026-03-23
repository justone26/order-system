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

# --- [4단계: 데이터 편집 및 재고 관리] ---
            st.divider()
            st.subheader("📊 4단계: 데이터 편집 및 재고 관리")
            
            # 0. 데이터 복사
            df_work = st.session_state.df_raw.copy()

            # 1. 상단 레이아웃 (검색, 품절필터, 제외옵션, 날짜)
            f_c1, f_c2, f_c3, f_c4 = st.columns([1.5, 1, 1, 1])
            search_q = f_c1.text_input("🔍 상품명 검색")
            filter_m = f_c2.selectbox("품절 필터", ["전체보기", "정상만", "품절만"], index=1)
            
            # [핵심] 러블리마켓(QT/BE) 상품 노출 여부 선택
            show_all_items = f_c3.checkbox("QT/BE 상품 포함하기", value=False, help="러블리마켓 등 QT/BE가 붙은 상품을 목록에 표시합니다.")
            
            hist_date_4 = f_c4.date_input("🗓️ 입고 기록 확인 날짜", datetime.now(), key="date_4")

            # 2. 특정 상품 제외 필터링 (체크박스가 해제되었을 때만 작동)
            if not show_all_items and item in df_work.columns:
                # 상품명에 'QT' 포함 혹은 'BE' 시작 상품 제외
                df_work = df_work[~df_work[item].astype(str).str.contains('QT|qt', na=False)]
                df_work = df_work[~df_work[item].astype(str).str.contains('^BE|^be', na=False)]

            # 3. 고유 키 생성 및 리오더 수량/입고 데이터 매핑
            df_work['unique_key'] = df_work[item].astype(str).str.strip() + df_work[option].astype(str).str.strip()
            
            # 과거 입고 수량 불러오기
            past_hist = load_history_from_gsheet()
            df_work['리오더입고수량'] = 0
            if not past_hist.empty and '구분' in past_hist.columns:
                past_hist['날짜'] = past_hist['저장시간'].astype(str).str.split(' ').str[0]
                t_hist = past_hist[(past_hist['날짜'] == hist_date_4.strftime("%Y-%m-%d")) & (past_hist['구분'] == "입고")]
                if not t_hist.empty:
                    t_hist['k_tmp'] = t_hist['상품명'].astype(str).str.strip() + t_hist['옵션'].astype(str).str.strip()
                    in_map = t_hist.groupby('k_tmp')['수량'].sum().to_dict()
                    df_work['리오더입고수량'] = df_work['unique_key'].map(in_map).fillna(0).astype(int)

            # 4. 수치 계산 로직 (안전한 numeric 변환 함수 사용)
            def safe_num(val):
                res = pd.to_numeric(val, errors='coerce')
                return res.fillna(0) if hasattr(res, 'fillna') else (0 if pd.isna(res) else res)

            v7_n = safe_num(df_work[t7day])
            v3_n = safe_num(df_work[t3day])
            df_work['일판매량'] = (v7_n / 7 if v7_n.sum() > 0 else v3_n / 3).round(0).astype(int)
            
            v_av_n = safe_num(df_work[avail])
            v_re_n = safe_num(df_work['리오더 수량'])
            df_work['권장발주량'] = ((df_work['일판매량'] * (lt + ss)) - (v_av_n + v_re_n)).clip(lower=0).astype(int)
            
            # 5. 사용자 추가 필터링 (품절 여부 및 검색어)
            if filter_m == "정상만": 
                df_work = df_work[~df_work[sold_out].astype(str).str.contains('품절', na=False)]
            elif filter_m == "품절만": 
                df_work = df_work[df_work[sold_out].astype(str).str.contains('품절', na=False)]
            
            if search_q: 
                df_work = df_work[df_work[item].astype(str).str.contains(search_q, case=False)]

            # 6. 화면 출력용 데이터 가공 (왼쪽 정렬 및 문자열 변환)
            if not df_work.empty:
                disp_cols_4 = [sold_out, vendor, item, option, vendor_item, stock, avail, "리오더 수량", "리오더입고수량", t3day, "일판매량", "권장발주량"]
                df_work_view = df_work[disp_cols_4].copy()
                
                # 숫자 컬럼을 텍스트로 변환하여 왼쪽 정렬 유도
                for c in df_work_view.columns:
                    df_work_view[c] = df_work_view[c].astype(str)
                
                config_4 = {c: st.column_config.TextColumn(c) for c in df_work_view.columns}

                # 수정 시 자동 실행 및 구글 시트 저장 함수
                def on_edit_4():
                    changes = st.session_state["main_editor_v4"]["edited_rows"]
                    for r_idx_str, change in changes.items():
                        idx = int(r_idx_str)
                        orig_idx = df_work.index[idx]
                        # 리오더 수량 직접 수정
                        if "리오더 수량" in change:
                            try: st.session_state.df_raw.at[orig_idx, "리오더 수량"] = int(change["리오더 수량"])
                            except: pass
                        # 입고 수량 입력 시 리오더 수량 차감 및 로그 기록
                        if "리오더입고수량" in change:
                            try:
                                in_qty = int(change["리오더입고수량"])
                                curr_re = st.session_state.df_raw.at[orig_idx, "리오더 수량"]
                                st.session_state.df_raw.at[orig_idx, "리오더 수량"] = max(0, curr_re - in_qty)
                                save_history_to_gsheet(pd.DataFrame([[df_work.at[orig_idx, item], df_work.at[orig_idx, option], in_qty]], columns=['상품명', '옵션', '수량']), log_type="입고")
                            except: pass
                    # 변경사항 시트에 즉시 저장
                    save_reorder_data(st.session_state.df_raw[[item, option, '리오더 수량']].rename(columns={item:'상품명', option:'옵션'}))
                    st.rerun()

                st.data_editor(
                    df_work_view, 
                    use_container_width=True, 
                    key="main_editor_v4", 
                    on_change=on_edit_4, 
                    column_config=config_4, 
                    hide_index=True
                )
            else:
                st.warning("⚠️ 표시할 데이터가 없습니다. 필터 설정(QT/BE 포함 여부 등)을 확인해주세요.")

            # --- [5단계: 최종 발주 리스트 요약] ---
            st.divider()
            st.subheader("📋 5단계: 최종 발주 리스트 요약")
            
            c5_1, c5_2 = st.columns([2, 1])
            s_filter = c5_1.selectbox("🎯 상태 필터", ["🚨긴급 + ⚠️주의 우선", "🚨 긴급만 보기", "✅ 정상 포함 전체보기"], index=0)
            hist_date_5 = c5_2.date_input("🗓️ 입고 기록 확인 날짜 (연동)", value=hist_date_4, key="date_5")

            to_order = df_work.copy()
            # 실시간 발주 계산 (기존 추가발주량 딕셔너리 연동)
            to_order['과거입고수량'] = to_order['리오더입고수량'] 
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

            # [핵심] 상태 우선순위 정렬 로직
            status_rank = {"🚨 긴급": 0, "⚠️ 주의": 1, "✅ 정상": 2}
            to_order['rank'] = to_order['상태'].map(status_rank)
            to_order = to_order.sort_values(by='rank').drop(columns=['rank'])

            # 필터 조건에 따른 데이터 추출
            if s_filter == "🚨긴급 + ⚠️주의 우선":
                df_final = to_order[to_order['상태'].isin(["🚨 긴급", "⚠️ 주의"]) | (to_order['권장발주량'] > 0)].copy()
            elif s_filter == "🚨 긴급만 보기":
                df_final = to_order[to_order['상태'] == "🚨 긴급"].copy()
            else:
                df_final = to_order.copy()

            if not df_final.empty:
                disp_final = ["상태", item, option, vendor, vendor_item, avail, "리오더 수량", "추가발주수량", "과거입고수량", "권장발주량", "최종발주량"]
                df_final_view = df_final[disp_final].copy()
                for c in df_final_view.columns: df_final_view[c] = df_final_view[c].astype(str)
                config_5 = {c: st.column_config.TextColumn(c) for c in df_final_view.columns}
                
                def on_edit_5():
                    edits = st.session_state["final_editor_v5"]["edited_rows"]
                    for r_idx_str, change in edits.items():
                        if "추가발주수량" in change:
                            r_key = df_final.iloc[int(r_idx_str)]['unique_key']
                            try: st.session_state.extra_order_dict[r_key] = int(change["추가발주수량"])
                            except: pass
                    st.rerun()

                st.data_editor(df_final_view, use_container_width=True, key="final_editor_v5", column_config=config_5, on_change=on_edit_5, hide_index=True)
                
                # 저장 및 다운로드 버튼 (기존 로직 유지)
                c_b1, c_b2 = st.columns(2)
                if c_b1.button("💾 구글 시트에 최종 발주 기록 저장", use_container_width=True):
                    # 문자를 숫자로 변환하여 저장
                    save_df = df_final.copy()
                    save_df['최종발주량'] = pd.to_numeric(save_df['최종발주량'], errors='coerce').fillna(0)
                    save_df = save_df[save_df['최종발주량'] > 0][[item, option, '최종발주량']]
                    if save_history_to_gsheet(save_df): st.success("✅ 저장 완료!")
                
                csv = df_final[disp_final].to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
                c_b2.download_button("📥 엑셀 다운로드", csv, f"최종발주서_{datetime.now().strftime('%m%d')}.csv", use_container_width=True)
                
# --- [🌙 탭 2: 동대문 사입 관리] ---
with tab2:
    st.subheader("🌙 동대문 사입 및 미납 관리")
    dong_file = st.file_uploader("동대문 주문 리스트 업로드", type=['xlsx', 'csv'], key="dong_tab_upload")
    if dong_file:
        if "last_file_name" not in st.session_state or st.session_state.last_file_name != dong_file.name:
            df = pd.read_excel(dong_file) if not dong_file.name.endswith('.csv') else pd.read_csv(dong_file)
            df.columns = df.columns.str.strip()
            # ... (이하 동대문 로직 동일)
            st.session_state.df_dong_current = df
            st.session_state.last_file_name = dong_file.name
        st.dataframe(st.session_state.df_dong_current, use_container_width=True)
