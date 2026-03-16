import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- [1. 함수 정의] ---
def get_sheet():
    creds_dict = dict(st.secrets["gcp_service_account"])
    scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    spreadsheet_key = "1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg"
    return client.open_by_key(spreadsheet_key)

def save_reorder_data(df):
    try:
        sheet = get_sheet().sheet1 # 첫 번째 시트(발주용)
        sheet.clear()
        sheet.update([df.columns.values.tolist()] + df.values.tolist())
    except Exception as e:
        st.error(f"구글 시트 저장 실패: {e}")

def save_history_to_gsheet(df):
    try:
        spreadsheet = get_sheet()
        # 'history' 탭 가져오기 (없으면 자동 생성)
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

# --- [2. 앱 설정 및 세션 초기화] ---
st.set_page_config(layout="wide", page_title="재고 관리 시스템")
st.title("📦 재고 관리 및 발주 시스템")

if 'analyzed' not in st.session_state: st.session_state.analyzed = False

# --- [3. 데이터 업로드] ---
st.subheader("📁 데이터 업로드")
if st.button("🔄 전체 데이터 초기화 및 재업로드"):
    st.session_state.clear()
    st.rerun()

uploaded_file = st.file_uploader("엑셀/CSV 파일을 선택하세요", type=['xlsx', 'xls', 'csv'])
st.divider()

if uploaded_file is not None:
    if 'df_raw' not in st.session_state or st.session_state.get('last_filename') != uploaded_file.name:
        st.session_state.df_raw = pd.read_excel(uploaded_file).loc[:, ~pd.read_excel(uploaded_file).columns.duplicated()]
        st.session_state.last_filename = uploaded_file.name
        st.session_state.analyzed = False
        st.rerun()

# --- [4. 메인 로직] ---
if st.session_state.get('df_raw') is not None:
    cols = st.session_state.df_raw.columns.tolist()

    # 1단계: 매핑 설정
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

    # 2~3단계: 분석 설정
    st.subheader("⚙️ 2~3단계: 분석 설정")
    lead_time = st.number_input("리드타임 (일)", value=10)
    safety_stock = st.number_input("안전재고 (일 수)", value=7)
    
    if st.button("🚀 분석 실행"):
        df = st.session_state.df_raw.copy()
        daily_avg = pd.to_numeric(df[t1week], errors='coerce').fillna(0) / 7
        df['권장 발주량'] = ((daily_avg * lead_time) + (daily_avg * safety_stock) - pd.to_numeric(df[avail], errors='coerce').fillna(0)).clip(lower=0).astype(int)
        st.session_state.df_raw = df
        st.session_state.analyzed = True
        st.rerun()

    # 분석이 완료된 경우에만 표시
    if st.session_state.analyzed:
        # --- 4단계: 데이터 편집 ---
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

# --- 5단계: 발주 리스트 요약 ---
        st.subheader("📋 5단계: 발주 리스트 요약")
        
        # 권장 발주량이 0보다 큰 항목만 추출
        to_order = st.session_state.df_raw[st.session_state.df_raw['권장 발주량'] > 0].copy()
        
        if not to_order.empty:
            # 초기값 설정
            if '추가 리오더' not in to_order.columns:
                to_order['추가 리오더'] = 0
            
            # 출력 컬럼 설정 (옵션 다음 공급처 상품명 배치)
            display_cols = [vendor, item, option, vendor_item, "리오더 수량", "추가 리오더", "권장 발주량"]
            
            # [1] 편집기 부분
            st.write("### ✏️ 발주 수량 입력") 
            edited = st.data_editor(to_order[display_cols], use_container_width=True, key="order_editor")
            
            # [2] 긴급 발주 확인 부분 (색상 강조)
            st.write("### 🚨 긴급 발주 상품 확인")
            def highlight_urgent(row):
                # 3일 평균 판매량 대비 가용재고가 적으면 행 전체 분홍색 (#ffcccc)
                avg_3d = float(row.get(t3day, 0)) / 3
                avail_val = float(row.get(avail, 0))
                return ['background-color: #ffcccc'] * len(row) if avail_val <= avg_3d else [''] * len(row)
            
            # 편집기(edited) 데이터를 실시간으로 스타일링하여 표시
            st.dataframe(edited.style.apply(highlight_urgent, axis=1), use_container_width=True)
            
            # [3] 저장 및 다운로드 버튼 영역
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button("💾 구글 시트 및 기록 저장"):
                    # 1. 원본 데이터프레임에 수정된 수량 반영 (리오더 수량 + 추가 리오더)
                    for idx, row in edited.iterrows():
                        st.session_state.df_raw.at[idx, '리오더 수량'] = float(row['리오더 수량']) + float(row['추가 리오더'])
                    
                    # 2. 메인 발주 현황 시트 업데이트 (덮어쓰기)
                    save_reorder_data(st.session_state.df_raw[[item, option, '리오더 수량']])
                    
                    # 3. 'history' 탭에 누적 저장 (영구 기록용)
                    if save_history_to_gsheet(edited):
                        st.success("✅ 구글 시트 및 히스토리가 영구 저장되었습니다!")
            
            with col2:
                # 엑셀 다운로드 버튼
                csv_data = edited.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="📥 발주 리스트 엑셀 다운로드",
                    data=csv_data,
                    file_name=f"발주_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv"
                )
        else:
            st.info("💡 권장 발주량이 있는 상품이 없습니다. 분석 단계를 먼저 확인해주세요.")
            
        # --- 6단계: 과거 데이터 확인 ---
        st.subheader("📜 6단계: 과거 데이터 확인")
        if st.button("🔄 구글 시트에서 기록 불러오기"):
            st.session_state.db_history = load_history_from_gsheet()

        if 'db_history' in st.session_state and not st.session_state.db_history.empty:
            df_hist = st.session_state.db_history
            df_hist['날짜'] = df_hist['저장시간'].astype(str).str.split(' ').str[0]
            selected_date = st.date_input("조회할 날짜 선택", datetime.now())
            target_date_str = selected_date.strftime("%Y-%m-%d")
            
            day_data = df_hist[df_hist['날짜'] == target_date_str]
            if not day_data.empty:
                selected_time = st.selectbox("⏰ 저장 시간 선택", sorted(day_data['저장시간'].unique(), reverse=True))
                final_view = day_data[day_data['저장시간'] == selected_time].drop(columns=['날짜'])
                st.dataframe(final_view, use_container_width=True)
            else:
                st.info(f"📅 {target_date_str} 기록이 없습니다.")
