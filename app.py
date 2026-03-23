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

# [핵심] History 시트에서 각 상품별로 최근 3개의 입고 기록을 가져오는 함수
def get_recent_in_logs():
    try:
        spreadsheet = get_sheet()
        hist_sheet = spreadsheet.worksheet("history")
        data = hist_sheet.get_all_records()
        if not data: return {}
        
        df_h = pd.DataFrame(data)
        # 입고 수량이 입력된 기록만 필터링 (컬럼명 확인 필요)
        if '리오더입고수량' not in df_h.columns: return {}
        
        df_in = df_h[pd.to_numeric(df_h['리오더입고수량'], errors='coerce') > 0].copy()
        if df_in.empty: return {}

        # 최신순 정렬
        df_in = df_in.sort_values(by='저장시간', ascending=False)
        
        log_dict = {}
        for (name, opt), group in df_in.groupby(['상품명', '옵션']):
            # 최근 3건만 추출해서 "날짜(수량)" 형태로 변환
            recent_rows = group.head(3)
            log_strings = [f"{str(r['저장시간'])[5:10]}({int(r['리오더입고수량'])}개)" for _, r in recent_rows.iterrows()]
            log_dict[(str(name).strip(), str(opt).strip())] = " / ".join(log_strings)
        return log_dict
    except:
        return {}

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
            hist_sheet = spreadsheet.add_worksheet(title="history", rows="1000", cols="25")
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

# --- [🏭 탭 1: 제작 상품 관리] ---
with tab1:
    if 'analyzed' not in st.session_state: st.session_state.analyzed = False
    
    st.subheader("📁 데이터 업로드 (제작상품)")
    
    if st.button("🔄 제작상품 분석 리셋"):
        for key in ['df_raw', 'analyzed', 'last_filename', 'in_logs']:
            if key in st.session_state: st.session_state.pop(key, None)
        st.rerun()

    uploaded_file = st.file_uploader("엑셀/CSV 파일을 선택하세요", type=['xlsx', 'xls', 'csv'], key="prod_upload")
    st.divider()

    if uploaded_file is not None:
        if 'df_raw' not in st.session_state or st.session_state.get('last_filename') != uploaded_file.name:
            df_new = pd.read_excel(uploaded_file)
            df_new.columns = df_new.columns.str.strip()
            df_new = df_new.loc[:, ~df_new.columns.duplicated()] 
            
            # 기본 상품명/옵션 매칭용 컬럼 찾기
            tmp_item = next((c for c in df_new.columns if '상품명' in c), df_new.columns[0])
            tmp_option = next((c for c in df_new.columns if '옵션' in c), df_new.columns[min(1, len(df_new.columns)-1)])

            try:
                # 1. 시트1에서 기존 리오더 수량 로드
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
                
                # 2. History에서 입고 내역 로그 로드
                st.session_state.in_logs = get_recent_in_logs()
                
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

        st.subheader("⚙️ 1~3단계: 설정 및 분석")
        c1, c2 = st.columns(2)
        sold_out = c1.selectbox("품절 여부", cols, index=find_idx(cols, ['품절']))
        item = c1.selectbox("상품명", cols, index=find_idx(cols, ['상품명']))
        option = c1.selectbox("옵션", cols, index=find_idx(cols, ['옵션']))
        avail = c1.selectbox("가용재고", cols, index=find_idx(cols, ['가용재고']))
        t3day = c2.selectbox("3일 발주합계", cols, index=find_idx(cols, ['3일']))
        lead_time = c2.number_input("리드타임 (일)", value=10)
        safety_stock = c2.number_input("안전재고 (일 수)", value=7)
        
        if st.button("🚀 분석 실행"):
            st.session_state.analyzed = True
            st.rerun()

        if st.session_state.analyzed:
            st.subheader("📊 4단계: 입고 관리 및 과거 내역 확인")
            
            df_working = st.session_state.df_raw.copy()
            
            # [기능 핵심] '최근입고내역' 열 추가
            def get_log_str(row):
                key = (str(row[item]).strip(), str(row[option]).strip())
                return st.session_state.get('in_logs', {}).get(key, "-")

            df_working['최근입고내역'] = df_working.apply(get_log_str, axis=1)

            # 계산 로직
            v_avail = pd.to_numeric(df_working[avail], errors='coerce').fillna(0)
            v_reorder = pd.to_numeric(df_working['리오더 수량'], errors='coerce').fillna(0)
            v_3day = pd.to_numeric(df_working[t3day], errors='coerce').fillna(0)
            df_working["일판매량"] = (v_3day / 3).round(0).astype(int)
            needed_qty = (df_working["일판매량"] * (lead_time + safety_stock))
            df_working["권장발주량"] = (needed_qty - (v_avail + v_reorder)).clip(lower=0).round(0).astype(int)
            
            # 품절(단종) 처리
            is_sold_out = df_working[sold_out].astype(str).str.contains('품절', na=False)
            df_working.loc[is_sold_out, "권장발주량"] = 0

            # 검색 및 필터
            f1, f2 = st.columns([3, 1])
            search_query = f1.text_input("🔍 상품명 검색")
            filter_mode = f2.selectbox("품절 필터", ["전체보기", "정상만", "품절만"], index=1)
            
            if filter_mode == "정상만": df_working = df_working[~is_sold_out]
            elif filter_mode == "품절만": df_working = df_working[is_sold_out]
            if search_query: df_working = df_working[df_working[item].astype(str).str.contains(search_query, case=False)]

            if "리오더입고수량" not in df_working.columns: df_working["리오더입고수량"] = 0

            # 데이터 에디터 저장 로직
            def on_edit_save():
                if "main_editor" in st.session_state and st.session_state["main_editor"]["edited_rows"]:
                    changes = st.session_state["main_editor"]["edited_rows"]
                    for row_idx_str, change in changes.items():
                        row_idx = int(row_idx_str)
                        orig_idx = df_working.index[row_idx]
                        
                        # 입고 수량 입력 시 차감 및 개별 기록
                        if "리오더입고수량" in change:
                            in_qty = int(float(change["리오더입고수량"]))
                            st.session_state.df_raw.at[orig_idx, "리오더 수량"] -= in_qty
                            # 개별 기록용 한 줄 저장
                            row_log = df_working.iloc[[row_idx]].copy()
                            row_log["리오더입고수량"] = in_qty
                            save_history_to_gsheet(row_log[[item, option, '리오더입고수량']])
                    
                    # 시트1 동기화
                    save_df = st.session_state.df_raw[[item, option, '리오더 수량']].copy()
                    save_df.columns = ['상품명', '옵션', '리오더 수량']
                    save_reorder_data(save_df)
                    # 내역 갱신을 위해 로그 재로드
                    st.session_state.in_logs = get_recent_in_logs()
                    st.rerun()

            # 화면 표시용 열 구성
            display_cols = [item, option, avail, "리오더 수량", "리오더입고수량", "최근입고내역", "권장발주량"]
            
            st.data_editor(
                df_working[display_cols], use_container_width=True, height=500, key="main_editor",
                on_change=on_edit_save,
                column_config={
                    "최근입고내역": st.column_config.Column("📅 최근 입고 기록 (날짜/수량)", width="large"),
                    "리오더입고수량": st.column_config.Column("➕ 입고 수량 입력"),
                    "리오더 수량": st.column_config.Column("📝 남은 리오더"),
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
