import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- [1. 공통 함수 정의] ---
def get_sheet():
    creds_dict = dict(st.secrets["gcp_service_account"])
    scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    spreadsheet_key = "1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg"
    return client.open_by_key(spreadsheet_key)

# [추가] History에서 품목별 마지막 입고 기록을 가져오는 함수
def get_last_history_data():
    try:
        spreadsheet = get_sheet()
        hist_sheet = spreadsheet.worksheet("history")
        data = hist_sheet.get_all_records()
        if not data: return pd.DataFrame()
        
        df_h = pd.DataFrame(data)
        # 입고 수량이 있었던 기록만 필터링 (추가발주분이나 리오더 수량이 0보다 큰 기록)
        # 여기서는 '저장시간' 기준으로 가장 최신 데이터를 가져옵니다.
        df_h = df_h.sort_values(by='저장시간', ascending=False)
        # 상품명+옵션 조합으로 가장 최근 것만 남김
        df_last = df_h.drop_duplicates(subset=['상품명', '옵션'], keep='first')
        return df_last[['상품명', '옵션', '저장시간', '권장발주량']] # 권장발주량 혹은 입고량 열 이름에 맞게 수정 가능
    except:
        return pd.DataFrame()

def save_reorder_data(df):
    if df.empty: return 
    try:
        sheet = get_sheet().sheet1
        data = [df.columns.values.tolist()] + df.values.tolist()
        sheet.update('A1', data) 
    except Exception as e:
        st.error(f"구글 시트 저장 실패: {e}")

def save_history_to_gsheet(df):
    try:
        spreadsheet = get_sheet()
        try:
            hist_sheet = spreadsheet.worksheet("history")
        except:
            hist_sheet = spreadsheet.add_worksheet(title="history", rows="1000", cols="20")
            hist_sheet.append_row(["저장시간"] + df.columns.tolist())
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        rows_to_add = [[now_str] + row for row in df.values.tolist()]
        hist_sheet.append_rows(rows_to_add)
        return True
    except Exception as e:
        st.error(f"과거 기록 저장 실패: {e}")
        return False

def find_idx(cols, target_keywords):
    for keyword in target_keywords:
        for i, col in enumerate(cols):
            if keyword in str(col): return i
    return 0

# --- [2. 앱 설정 및 탭 구성] ---
st.set_page_config(layout="wide", page_title="재고 관리 시스템")
st.title("📦 통합 재고 관리 시스템")

tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 관리"])

with tab1:
    if 'analyzed' not in st.session_state: st.session_state.analyzed = False
    
    st.subheader("📁 데이터 업로드 (제작상품)")
    
    if st.button("🔄 제작상품 분석 리셋"):
        for key in ['df_raw', 'analyzed', 'last_filename', 'df_history_match']:
            if key in st.session_state: del st.session_state[key]
        st.rerun()

    uploaded_file = st.file_uploader("엑셀/CSV 파일을 선택하세요", type=['xlsx', 'xls', 'csv'], key="prod_upload")
    st.divider()

    if uploaded_file is not None:
        if 'df_raw' not in st.session_state or st.session_state.get('last_filename') != uploaded_file.name:
            df_new = pd.read_excel(uploaded_file)
            df_new.columns = df_new.columns.str.strip()
            df_new = df_new.loc[:, ~df_new.columns.duplicated()] 
            
            tmp_item = next((c for c in df_new.columns if '상품명' in c), df_new.columns[0])
            tmp_option = next((c for c in df_new.columns if '옵션' in c), df_new.columns[min(1, len(df_new.columns)-1)])

            try:
                sheet = get_sheet().sheet1
                gs_raw = sheet.get_all_records()
                if gs_raw:
                    gs_data = pd.DataFrame(gs_raw)
                    if not gs_data.empty and '상품명' in gs_data.columns:
                        gs_subset = gs_data[['상품명', '옵션', '리오더 수량']].copy()
                        df_new['match_name'] = df_new[tmp_item].astype(str).str.strip()
                        df_new['match_opt'] = df_new[tmp_option].astype(str).str.strip()
                        df_new = pd.merge(df_new, gs_subset, left_on=['match_name', 'match_opt'], right_on=['상품명', '옵션'], how='left', suffixes=('', '_gs'))
                        if '리오더 수량_gs' in df_new.columns:
                            df_new['리오더 수량'] = df_new['리오더 수량_gs'].fillna(0).astype(int)
                            cols_to_drop = [c for c in ['상품명_gs', '옵션_gs', '리오더 수량_gs', 'match_name', 'match_opt'] if c in df_new.columns]
                            df_new = df_new.drop(columns=cols_to_drop)
                
                # [추가] History 데이터 미리 불러오기
                st.session_state.df_history_match = get_last_history_data()

                if '리오더 수량' not in df_new.columns: df_new['리오더 수량'] = 0
            except Exception as e:
                st.warning(f"데이터 연동 실패: {e}")

            st.session_state.df_raw = df_new
            st.session_state.last_filename = uploaded_file.name
            st.session_state.analyzed = False
            st.rerun()

    if st.session_state.get('df_raw') is not None:
        df_current = st.session_state.df_raw
        cols = df_current.columns.tolist()

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

        st.subheader("⚙️ 2~3단계: 분석 설정")
        col_lt, col_ss = st.columns(2)
        lead_time = col_lt.number_input("리드타임 (일)", value=10)
        safety_stock = col_ss.number_input("안전재고 (일 수)", value=7)
        
        if st.button("🚀 분석 실행"):
            st.session_state.analyzed = True
            st.rerun()

        if st.session_state.analyzed:
            st.subheader("📊 4단계: 데이터 편집 및 재고 관리")
            
            df_working = st.session_state.df_raw.copy()
            
            # 히스토리 데이터 매칭 (최근 입고일자 가져오기)
            if 'df_history_match' in st.session_state and not st.session_state.df_history_match.empty:
                hist = st.session_state.df_history_match
                df_working = pd.merge(df_working, hist, left_on=[item, option], right_on=['상품명', '옵션'], how='left', suffixes=('', '_hist'))
                df_working['최근기록일'] = df_working['저장시간'].fillna('-')
            else:
                df_working['최근기록일'] = '-'

            v_avail = pd.to_numeric(df_working[avail], errors='coerce').fillna(0)
            v_reorder = pd.to_numeric(df_working['리오더 수량'], errors='coerce').fillna(0)
            v_3day = pd.to_numeric(df_working[t3day], errors='coerce').fillna(0)
            df_working["일판매량"] = (v_3day / 3).round(0).astype(int)
            
            needed_qty = (df_working["일판매량"] * (lead_time + safety_stock))
            current_assets = (v_avail + v_reorder)
            df_working["권장발주량"] = (needed_qty - current_assets).clip(lower=0).round(0).astype(int)
            
            is_sold_out = df_working[sold_out].astype(str).str.contains('품절', na=False)
            df_working.loc[is_sold_out, "권장발주량"] = 0

            # 검색 및 필터
            f1, f2 = st.columns([3, 1])
            search_query = f1.text_input("🔍 상품명 검색", key="prod_search_input")
            filter_mode = f2.selectbox("품절 필터", ["전체보기", "정상만", "품절만"], index=1)
            
            if filter_mode == "정상만": df_working = df_working[~is_sold_out]
            elif filter_mode == "품절만": df_working = df_working[is_sold_out]
            if search_query: df_working = df_working[df_working[item].astype(str).str.contains(search_query, case=False, na=False)]

            if "리오더입고수량" not in df_working.columns: df_working["리오더입고수량"] = 0

            # 자동 저장 로직
            def auto_save_and_update():
                if "main_editor" in st.session_state and st.session_state["main_editor"]["edited_rows"]:
                    changes = st.session_state["main_editor"]["edited_rows"]
                    for row_idx_str, change in changes.items():
                        row_idx = int(row_idx_str)
                        orig_idx = df_working.index[row_idx]
                        if "리오더 수량" in change:
                            val = str(change["리오더 수량"]).strip()
                            st.session_state.df_raw.at[orig_idx, "리오더 수량"] = int(float(val)) if val else 0
                        if "리오더입고수량" in change:
                            val = str(change["리오더입고수량"]).strip()
                            in_qty = int(float(val)) if val else 0
                            curr = st.session_state.df_raw.at[orig_idx, "리오더 수량"]
                            st.session_state.df_raw.at[orig_idx, "리오더 수량"] = max(0, curr - in_qty)
                    try:
                        save_df = st.session_state.df_raw[[item, option, '리오더 수량']].copy()
                        save_df.columns = ['상품명', '옵션', '리오더 수량']
                        save_reorder_data(save_df)
                        st.toast("✅ 수량 업데이트 완료!")
                    except: pass

            display_df_4 = df_working.copy()
            # 숫자 데이터 정리
            for col in [stock, avail, "리오더 수량", "리오더입고수량", "권장발주량"]:
                if col in display_df_4.columns:
                    display_df_4[col] = display_df_4[col].fillna(0).astype(int).apply(lambda x: f"  {x}")

            # [핵심] 최근기록일 열을 앞쪽에 배치
            display_cols_4 = [sold_out, "최근기록일", item, option, stock, avail, "리오더 수량", "리오더입고수량", "권장발주량"]
            final_target_4 = [c for c in display_cols_4 if c in display_df_4.columns]

            st.data_editor(
                display_df_4[final_target_4], use_container_width=True, height=450, key="main_editor",
                on_change=auto_save_and_update,
                column_config={
                    "최근기록일": st.column_config.Column("📅 마지막 저장일", width="medium"),
                    "리오더 수량": st.column_config.Column("📝 리오더 수량"),
                    "리오더입고수량": st.column_config.Column("➕ 입고수량 입력")
                }
            )

            st.subheader("📋 5단계: 최종 발주 리스트 요약")
            # 5단계에서도 동일하게 품절 상품은 제외하거나 수량을 0으로 처리
            to_order = st.session_state.df_raw.copy()
            v_3day_val = pd.to_numeric(to_order[t3day], errors='coerce').fillna(0)
            to_order['일판매량'] = (v_3day_val / 3).round(0).astype(int)
            
            # 분석 수량 다시 계산 (품절 고려)
            v_av_5 = pd.to_numeric(to_order[avail], errors='coerce').fillna(0)
            v_re_5 = pd.to_numeric(to_order['리오더 수량'], errors='coerce').fillna(0)
            to_order['권장발주량'] = ((to_order['일판매량'] * (lead_time + safety_stock)) - (v_av_5 + v_re_5)).clip(lower=0).round(0).astype(int)
            
            # 품절 상품은 권장발주량 0
            to_order.loc[to_order[sold_out].astype(str).str.contains('품절', na=False), '권장발주량'] = 0

            def check_urgency(row):
                # 품절이면 상태 체크 의미 없음
                if '품절' in str(row.get(sold_out, "")): return "⏹️ 단종"
                v_av = pd.to_numeric(row.get(avail, 0), errors='coerce') or 0
                v_sl = pd.to_numeric(row.get('일판매량', 0), errors='coerce') or 0
                v_re = pd.to_numeric(row.get('리오더 수량', 0), errors='coerce') or 0
                if v_sl > 0 and (v_av + v_re) < (v_sl * 3): return "🚨 긴급"
                elif v_sl > 0 and (v_av + v_re) < (v_sl * 5): return "⚠️ 주의"
                return "✅ 정상"

            to_order['상태'] = to_order.apply(check_urgency, axis=1)
            status_filter = st.selectbox("🎯 상태 필터", ["전체보기", "🚨 긴급만 보기", "⚠️ 주의이상 보기"], key="f_filter_v11")
            
            # 발주할 게 있거나 상태가 정상이 아닌 것들만 표시 (단, 단종은 제외)
            mask = (pd.to_numeric(to_order['권장발주량'], errors='coerce') > 0) & (to_order['상태'] != "⏹️ 단종")
            to_order = to_order[mask].copy()
            
            if "🚨" in status_filter: to_order = to_order[to_order['상태'] == "🚨 긴급"]
            elif "⚠️" in status_filter: to_order = to_order[to_order['상태'].str.contains("🚨|⚠️")]

            if not to_order.empty:
                display_5step = to_order.copy()
                if '추가발주분' not in display_5step.columns: display_5step['추가발주분'] = 0
                for col in [avail, "리오더 수량", "추가발주분", "권장발주량"]:
                    display_5step[col] = display_5step[col].fillna(0).astype(int).apply(lambda x: f"  {x}")

                final_display_cols = ["상태", item, option, vendor_item, avail, "리오더 수량", "추가발주분", "권장발주량"]
                existing_cols_5 = [c for c in final_display_cols if c in display_5step.columns]
                
                final_order_df = st.data_editor(
                    display_5step[existing_cols_5], use_container_width=True, height=400, key="final_order_editor_v11"
                )
                
                st.divider()
                col_btn1, col_btn2 = st.columns(2)
                if col_btn1.button("💾 구글 시트에 최종 기록 저장"):
                    save_data = final_order_df.copy()
                    for c in save_data.columns:
                        if save_data[c].dtype == object: save_data[c] = save_data[c].astype(str).str.strip()
                    if save_history_to_gsheet(save_data): st.success("✅ 저장 완료!")

                csv_data = final_order_df.copy()
                for c in csv_data.columns:
                    if csv_data[c].dtype == object: csv_data[c] = csv_data[c].astype(str).str.strip()
                csv = csv_data.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
                col_btn2.download_button("📥 엑셀 다운로드", csv, f"발주_{datetime.now().strftime('%m%d_%H%M')}.csv", "text/csv")
            else:
                st.info("💡 발주할 상품이 없거나 모든 부족 상품이 '품절(단종)' 상태입니다.")

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
