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

def save_reorder_data(df):
    try:
        sheet = get_sheet().sheet1
        sheet.clear()
        sheet.update([df.columns.values.tolist()] + df.values.tolist())
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

def load_history_from_gsheet():
    try:
        spreadsheet = get_sheet()
        hist_sheet = spreadsheet.worksheet("history")
        data = hist_sheet.get_all_records()
        return pd.DataFrame(data)
    except Exception as e:
        st.error(f"기록 불러오기 실패: {e}")
        return pd.DataFrame()

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
    if st.button("🔄 제작상품 데이터 초기화"):
        st.session_state.clear()
        st.rerun()

    uploaded_file = st.file_uploader("엑셀/CSV 파일을 선택하세요", type=['xlsx', 'xls', 'csv'], key="prod_upload")
    st.divider()

    # [수정된 부분] 파일 업로드 시 구글 시트 데이터를 병합함
    if uploaded_file is not None:
        if 'df_raw' not in st.session_state or st.session_state.get('last_filename') != uploaded_file.name:
            # 1. 업로드한 엑셀 읽기
            df_new = pd.read_excel(uploaded_file)
            df_new.columns = df_new.columns.str.strip()
            df_new = df_new.loc[:, ~df_new.columns.duplicated()] 
            
            # 2. 구글 시트에서 기존 리오더 수량 가져오기
            try:
                sheet = get_sheet().sheet1
                gs_data = pd.DataFrame(sheet.get_all_records())
                
                if not gs_data.empty and '상품명' in gs_data.columns and '리오더 수량' in gs_data.columns:
                    # 필요한 컬럼만 추출 (상품명, 옵션을 기준으로 매칭)
                    gs_data = gs_data[['상품명', '옵션', '리오더 수량']].copy()
                    # 엑셀 데이터와 구글 시트 데이터 합치기 (왼쪽 조인)
                    df_new = pd.merge(df_new, gs_data, on=['상품명', '옵션'], how='left', suffixes=('', '_gs'))
                    
                    # 구글 시트에 값이 있으면 그 값을 쓰고, 없으면 0 처리
                    if '리오더 수량_gs' in df_new.columns:
                        df_new['리오더 수량'] = df_new['리오더 수량_gs'].fillna(0)
                        df_new = df_new.drop(columns=['리오더 수량_gs'])
                else:
                    df_new['리오더 수량'] = 0
            except Exception as e:
                st.warning(f"구글 시트 연동 실패(새 데이터로 시작): {e}")
                df_new['리오더 수량'] = 0

            st.session_state.df_raw = df_new
            st.session_state.last_filename = uploaded_file.name
            st.session_state.analyzed = False
            st.rerun()

        # 1단계: 매핑
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

        # 2~3단계: 분석
        st.subheader("⚙️ 2~3단계: 분석 설정")
        col_lt, col_ss = st.columns(2)
        lead_time = col_lt.number_input("리드타임 (일)", value=10)
        safety_stock = col_ss.number_input("안전재고 (일 수)", value=7)
        
        if st.button("🚀 분석 실행"):
            df = st.session_state.df_raw.copy()
            daily_avg = pd.to_numeric(df[t1week], errors='coerce').fillna(0) / 7
            df['권장 발주량'] = ((daily_avg * lead_time) + (daily_avg * safety_stock) - pd.to_numeric(df[avail], errors='coerce').fillna(0)).clip(lower=0).astype(int)
            st.session_state.df_raw = df
            st.session_state.analyzed = True
            st.rerun()

        if st.session_state.analyzed:
            # 4단계: 편집
            st.subheader("📊 4단계: 데이터 편집")
            target_cols = [sold_out, vendor, item, option, vendor_item, "정상재고", "가용재고", "리오더 수량", "리오더입고수량", "일판매량", "3일발주합계", "1주발주합계", "권장 발주량"]
            
            if "3일발주합계" in st.session_state.df_raw.columns:
                st.session_state.df_raw["일판매량"] = (pd.to_numeric(st.session_state.df_raw[t3day], errors='coerce').fillna(0) / 3).round(0).astype(int)

            for c in target_cols:
                if c not in st.session_state.df_raw.columns: st.session_state.df_raw[c] = 0

            f1, f2 = st.columns([3, 1])
            search_query = f1.text_input("🔍 상품명 검색")
            filter_mode = f2.selectbox("품절 필터", ["전체보기", "정상만", "품절만"], index=1)
            
            df_working = st.session_state.df_raw.copy()
            if filter_mode == "정상만": df_working = df_working[~df_working[sold_out].astype(str).str.contains('품절', na=False)]
            elif filter_mode == "품절만": df_working = df_working[df_working[sold_out].astype(str).str.contains('품절', na=False)]
            if search_query: df_working = df_working[df_working[item].astype(str).str.contains(search_query, case=False, na=False)]

            def update_reorder():
                edited = st.session_state["main_editor"]
                for row_idx, changes in edited['edited_rows'].items():
                    if '리오더입고수량' in changes:
                        received = float(changes['리오더입고수량'])
                        original_idx = df_working.index[row_idx]
                        st.session_state.df_raw.at[original_idx, '리오더 수량'] = max(0, float(st.session_state.df_raw.at[original_idx, '리오더 수량']) - received)
                        st.session_state.df_raw.at[original_idx, '리오더입고수량'] = 0

            st.data_editor(df_working[target_cols], use_container_width=True, key="main_editor", on_change=update_reorder)

            # 5단계: 요약
            st.subheader("📋 5단계: 발주 리스트 요약")
            to_order = st.session_state.df_raw[st.session_state.df_raw['권장 발주량'] > 0].copy()
            if not to_order.empty:
                to_order['추가 리오더'] = 0
                display_cols = [vendor, item, option, vendor_item, "리오더 수량", "추가 리오더", "권장 발주량"]
                st.write("### ✏️ 발주 수량 입력")
                edited = st.data_editor(to_order[display_cols], use_container_width=True, key="order_editor")
                st.write("### 🚨 긴급 발주 상품 확인")
                def highlight_urgent(row):
                    avg_3d = float(row.get(t3day, 0)) / 3
                    avail_val = float(row.get(avail, 0))
                    return ['background-color: #ffcccc'] * len(row) if avail_val <= avg_3d else [''] * len(row)
                st.dataframe(edited.style.apply(highlight_urgent, axis=1), use_container_width=True)
                
                c_s1, c_s2 = st.columns(2)
                if c_s1.button("💾 구글 시트 및 기록 저장"):
                    for idx, row in edited.iterrows():
                        st.session_state.df_raw.at[idx, '리오더 수량'] = float(row['리오더 수량']) + float(row['추가 리오더'])
                    save_reorder_data(st.session_state.df_raw[[item, option, '리오더 수량']])
                    if save_history_to_gsheet(edited): st.success("✅ 히스토리 저장 완료!")
                with c_s2:
                    csv_data = edited.to_csv(index=False).encode('utf-8-sig')
                    st.download_button("📥 엑셀 다운로드", csv_data, f"발주_{datetime.now().strftime('%Y%m%d')}.csv")

            # 6단계: 과거 확인
            st.subheader("📜 6단계: 과거 데이터 확인")
            if st.button("🔄 기록 불러오기"):
                st.session_state.db_history = load_history_from_gsheet()
            if 'db_history' in st.session_state and not st.session_state.db_history.empty:
                df_hist = st.session_state.db_history
                df_hist['날짜'] = df_hist['저장시간'].astype(str).str.split(' ').str[0]
                sel_date = st.date_input("날짜 선택", datetime.now())
                target_date = sel_date.strftime("%Y-%m-%d")
                day_data = df_hist[df_hist['날짜'] == target_date]
                if not day_data.empty:
                    sel_time = st.selectbox("⏰ 시간 선택", sorted(day_data['저장시간'].unique(), reverse=True))
                    st.dataframe(day_data[day_data['저장시간'] == sel_time].drop(columns=['날짜']), use_container_width=True)
                else: st.info("📅 해당 날짜 기록 없음")

# --- [🌙 탭 2: 동대문 사입 관리] ---
with tab2:
    st.subheader("🌙 동대문 사입 및 미납 관리")
    dong_file = st.file_uploader("동대문 주문 리스트 업로드", type=['xlsx', 'csv'], key="dong_tab_upload")
    
    # 1. 파일 업로드 및 데이터 처리
    if dong_file:
        if "last_file_name" not in st.session_state or st.session_state.last_file_name != dong_file.name:
            df = pd.read_excel(dong_file)
            df.columns = df.columns.str.strip()
            
            # [에러 방지] 엑셀에 컬럼이 없어도 강제로 생성
            required_cols = ['선택', '품절', '상품명', '공급처', '공급처상품명', '정상재고', '가용재고', '판매수량', '발주수량', '가중율', '3일판매']
            for col in required_cols:
                if col not in df.columns:
                    df[col] = 0 if col not in ['선택', '품절', '상품명', '공급처', '공급처상품명'] else ""
            
            # 수치형 변환
            for col in ['정상재고', '가용재고', '3일판매']:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
            # 계산 로직
            df['판매수량'] = (df['정상재고'] - df['가용재고']).clip(lower=0)
            df['가중율'] = df['판매수량'].apply(lambda n: 2.0 if n >= 10 else (1.5 if n >= 6 else (1.2 if n >= 3 else 1.0)))
            df['발주수량'] = (df['판매수량'] * df['가중율']).astype(int)
            
            st.session_state.df_dong_current = df[required_cols]
            st.session_state.last_file_name = dong_file.name

        # 2. 화면 출력
        df_display = st.session_state.df_dong_current.copy()
        
        # [검색창]
        c1, c2 = st.columns([1, 2])
        search_query = c2.text_input("상품명 검색")
        if search_query:
            df_display = df_display[df_display['상품명'].astype(str).str.contains(search_query, case=False, na=False)]

        # [데이터 편집기]
        df_display['선택'] = df_display['선택'].astype(bool)
        edited_df = st.data_editor(
            df_display, use_container_width=True, key="final_editor",
            column_config={"선택": st.column_config.CheckboxColumn("선택", width="small")}
        )

        # [버튼]
        st.divider()
        col1, col2, col3 = st.columns(3)
        add_val = col1.number_input("추가 수량", value=1, min_value=1)
        
        if col2.button("🚀 선택한 상품 수량 더하기"):
            # 선택된 인덱스만 발주수량 업데이트
            selected = edited_df[edited_df['선택'] == True].index
            for idx in selected:
                st.session_state.df_dong_current.at[idx, '발주수량'] += add_val
            st.rerun()
            
        csv = edited_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
        col3.download_button("📥 엑셀 다운로드", csv, "사입리스트.csv", "text/csv")
