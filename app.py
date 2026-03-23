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

# History에서 입고 내역 가져오기
def get_recent_in_logs():
    try:
        spreadsheet = get_sheet()
        hist_sheet = spreadsheet.worksheet("history")
        data = hist_sheet.get_all_records()
        if not data: return {}
        df_h = pd.DataFrame(data)
        if '리오더입고수량' not in df_h.columns: return {}
        df_in = df_h[pd.to_numeric(df_h['리오더입고수량'], errors='coerce') > 0].copy()
        if df_in.empty: return {}
        df_in = df_in.sort_values(by='저장시간', ascending=False)
        log_dict = {}
        for (name, opt), group in df_in.groupby(['상품명', '옵션']):
            recent_rows = group.head(3)
            log_strings = [f"{str(r['저장시간'])[5:10]}({int(r['리오더입고수량'])}개)" for _, r in recent_rows.iterrows()]
            log_dict[(str(name).strip(), str(opt).strip())] = " / ".join(log_strings)
        return log_dict
    except: return {}

def save_reorder_data(df):
    if df.empty: return 
    try:
        sheet = get_sheet().sheet1
        data = [df.columns.values.tolist()] + df.values.tolist()
        sheet.update('A1', data) 
    except Exception as e: st.error(f"구글 시트 저장 실패: {e}")

def save_history_to_gsheet(df):
    try:
        spreadsheet = get_sheet()
        try: hist_sheet = spreadsheet.worksheet("history")
        except:
            hist_sheet = spreadsheet.add_worksheet(title="history", rows="1000", cols="25")
            hist_sheet.append_row(["저장시간"] + df.columns.tolist())
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        rows_to_add = [[now_str] + [str(v) for v in row] for row in df.values.tolist()]
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

# --- [2. 앱 설정] ---
st.set_page_config(layout="wide", page_title="재고 관리 시스템")
st.title("📦 통합 재고 관리 시스템")

tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 관리"])

with tab1:
    if 'analyzed' not in st.session_state: st.session_state.analyzed = False
    
    st.subheader("📁 데이터 업로드 (제작상품)")
    if st.button("🔄 리셋"):
        for key in ['df_raw', 'analyzed', 'last_filename', 'in_logs']:
            st.session_state.pop(key, None)
        st.rerun()

    uploaded_file = st.file_uploader("엑셀 파일을 선택하세요", type=['xlsx', 'xls', 'csv'])

    if uploaded_file:
        if 'df_raw' not in st.session_state or st.session_state.get('last_filename') != uploaded_file.name:
            df_new = pd.read_excel(uploaded_file)
            df_new.columns = df_new.columns.str.strip()
            # 초기 리오더 수량 매칭
            try:
                st.session_state.in_logs = get_recent_in_logs()
                if '리오더 수량' not in df_new.columns: df_new['리오더 수량'] = 0
            except: pass
            st.session_state.df_raw = df_new
            st.session_state.last_filename = uploaded_file.name
            st.rerun()

    if st.session_state.get('df_raw') is not None:
        cols = st.session_state.df_raw.columns.tolist()
        st.subheader("⚙️ 분석 설정")
        c1, c2 = st.columns(2)
        sold_out = c1.selectbox("품절 여부", cols, index=find_idx(cols, ['품절']))
        item = c1.selectbox("상품명", cols, index=find_idx(cols, ['상품명']))
        option = c1.selectbox("옵션", cols, index=find_idx(cols, ['옵션']))
        vendor_item = c1.selectbox("공급처 상품명", cols, index=find_idx(cols, ['공급처상품명']))
        avail = c1.selectbox("가용재고", cols, index=find_idx(cols, ['가용재고']))
        t3day = c2.selectbox("3일 발주합계", cols, index=find_idx(cols, ['3일']))
        lead_time = c2.number_input("리드타임 (일)", value=10)
        safety_stock = c2.number_input("안전재고 (일 수)", value=7)
        
        if st.button("🚀 분석 실행"):
            st.session_state.analyzed = True
            st.rerun()

        if st.session_state.analyzed:
            st.subheader("📊 4단계: 입고 관리 및 내역 확인")
            df_working = st.session_state.df_raw.copy()
            
            # 기록 매칭
            def get_log_str(row):
                key = (str(row[item]).strip(), str(row[option]).strip())
                return st.session_state.get('in_logs', {}).get(key, "-")
            df_working['최근입고내역'] = df_working.apply(get_log_str, axis=1)

            # 계산
            v_avail = pd.to_numeric(df_working[avail], errors='coerce').fillna(0)
            v_reorder = pd.to_numeric(df_working['리오더 수량'], errors='coerce').fillna(0)
            v_3day = pd.to_numeric(df_working[t3day], errors='coerce').fillna(0)
            df_working["일판매량"] = (v_3day / 3).round(0).astype(int)
            df_working["권장발주량"] = ((df_working["일판매량"] * (lead_time + safety_stock)) - (v_avail + v_reorder)).clip(lower=0).round(0).astype(int)
            df_working.loc[df_working[sold_out].astype(str).str.contains('품절', na=False), "권장발주량"] = 0

            if "리오더입고수량" not in df_working.columns: df_working["리오더입고수량"] = 0

            # 4단계 자동 저장/기록 에디터
            def on_edit_4():
                if "main_editor" in st.session_state and st.session_state["main_editor"]["edited_rows"]:
                    changes = st.session_state["main_editor"]["edited_rows"]
                    for r_idx_str, change in changes.items():
                        r_idx = int(r_idx_str)
                        orig_idx = df_working.index[r_idx]
                        if "리오더입고수량" in change:
                            in_qty = int(change["리오더입고수량"])
                            st.session_state.df_raw.at[orig_idx, "리오더 수량"] -= in_qty
                            # 입고 개별 로그 저장
                            log_row = df_working.iloc[[r_idx]].copy()
                            log_row["리오더입고수량"] = in_qty
                            save_history_to_gsheet(log_row[[item, option, '리오더입고수량']])
                    
                    # 시트1 동기화
                    save_df = st.session_state.df_raw[[item, option, '리오더 수량']].copy()
                    save_df.columns = ['상품명', '옵션', '리오더 수량']
                    save_reorder_data(save_df)
                    st.session_state.in_logs = get_recent_in_logs()
                    st.rerun()

            display_cols_4 = [item, option, avail, "리오더 수량", "리오더입고수량", "최근입고내역", "권장발주량"]
            st.data_editor(df_working[display_cols_4], use_container_width=True, height=400, key="main_editor", on_change=on_edit_4)

            # --- [5단계: 최종 발주 요약 및 저장/다운로드] ---
            st.divider()
            st.subheader("📋 5단계: 최종 발주 리스트")
            
            # 발주 필요한 품목만 필터링
            df_to_order = df_working[df_working["권장발주량"] > 0].copy()
            
            if not df_to_order.empty:
                final_cols = [item, option, vendor_item, avail, "리오더 수량", "권장발주량"]
                # 5단계 편집기 (추가 수량 조절 가능하게 하려면)
                final_order_df = st.data_editor(df_to_order[final_cols], use_container_width=True, key="final_order_editor")
                
                col_save, col_down = st.columns(2)
                
                # 버튼 1: 구글 시트 저장
                if col_save.button("💾 구글 시트(History)에 최종 기록 저장"):
                    if save_history_to_gsheet(final_order_df):
                        st.success("✅ 발주 리스트가 성공적으로 기록되었습니다!")
                
                # 버튼 2: 엑셀 다운로드
                csv = final_order_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
                col_down.download_button(
                    label="📥 업체용 엑셀 다운로드",
                    data=csv,
                    file_name=f"발주리스트_{datetime.now().strftime('%m%d_%H%M')}.csv",
                    mime="text/csv"
                )
            else:
                st.info("💡 권장 발주량이 있는 상품이 없습니다.")

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
