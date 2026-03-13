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
    return client.open_by_key(spreadsheet_key).sheet1

def save_reorder_data(df):
    try:
        sheet = get_sheet()
        sheet.clear()
        sheet.update([df.columns.values.tolist()] + df.values.tolist())
    except Exception as e:
        st.error(f"구글 시트 저장 실패: {e}")

def find_idx(cols, target_keywords):
    for keyword in target_keywords:
        for i, col in enumerate(cols):
            if keyword in str(col): return i
    return 0

# --- [2. 앱 설정 및 세션 초기화] ---
st.set_page_config(layout="wide", page_title="재고 관리 시스템")
st.title("📦 재고 관리 및 발주 시스템")
if 'history' not in st.session_state: st.session_state.history = {}

# --- [3. 데이터 업로드 및 초기화] ---
st.subheader("📁 데이터 업로드")
if st.button("🔄 전체 데이터 초기화 및 재업로드"):
    st.session_state.clear()
    st.rerun()

uploaded_file = st.file_uploader("엑셀/CSV 파일을 여기에 드래그하거나 선택하세요", type=['xlsx', 'xls', 'csv'])
st.divider()

if uploaded_file is not None:
    if 'last_filename' not in st.session_state or st.session_state.last_filename != uploaded_file.name:
        st.session_state.df_raw = pd.read_excel(uploaded_file).loc[:, ~pd.read_excel(uploaded_file).columns.duplicated()]
        st.session_state.last_filename = uploaded_file.name
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
        st.rerun()

   # 4단계: 데이터 편집 (상세 정보 포함)
    st.subheader("📊 4단계: 데이터 편집")
    
    # 1. 계산 로직 재확인 (분석 후 데이터가 없을 경우 대비)
    if '일판매량' not in st.session_state.df_raw.columns:
        st.session_state.df_raw['일판매량'] = (pd.to_numeric(st.session_state.df_raw.get(t3day, 0), errors='coerce') / 3).round(1)

    # 2. 품절 여부 필터 (사용자 요청 반영)
    is_filter = st.checkbox("품절 상품 제외하고 보기")
    df_to_edit = st.session_state.df_raw.copy()
    if is_filter:
        # 품절 여부 컬럼 값이 '품절'이 아닌 것만 필터링
        df_to_edit = df_to_edit[df_to_edit[sold_out] != '품절']

    # 3. 편집할 컬럼 순서 (사용자님이 말씀하신 항목들 모두 포함)
    display_cols = [
        sold_out, vendor, item, option, vendor_item, reg_date, 
        stock, avail, "리오더 수량", "리오더입고수량", t3day, t1week, 
        "일판매량", "권장 발주량"
    ]
    
    # 일부 컬럼이 엑셀에 없을 경우를 대비해 존재하는 컬럼만 필터링
    existing_cols = [c for c in display_cols if c in df_to_edit.columns]

    # 4. 데이터 편집기 실행 및 결과 저장
    edited_df = st.data_editor(
        df_to_edit[existing_cols], 
        use_container_width=True, 
        key="main_editor"
    )
    
    # 편집된 내용을 전체 데이터에 반영
    st.session_state.df_raw.update(edited_df)

    # 5단계: 발주 리스트 요약
    st.subheader("📋 5단계: 발주 리스트 요약")
    if '권장 발주량' in st.session_state.df_raw.columns:
        to_order = st.session_state.df_raw[st.session_state.df_raw['권장 발주량'] > 0].copy()
        if not to_order.empty:
            to_order['추가 리오더'] = 0
            display_cols = [vendor, item, option, "리오더 수량", "추가 리오더", "권장 발주량"]
            edited = st.data_editor(to_order[display_cols], use_container_width=True, key="order_editor")
            
            # 스타일링
            def highlight_urgent(row):
                avail_val = float(row.get('가용재고', 0))
                sale_3d = float(row.get(t3day, 0)) / 3
                return ['background-color: #ffcccc'] * len(row) if avail_val <= sale_3d else [''] * len(row)
            
            st.dataframe(edited.style.apply(highlight_urgent, axis=1), use_container_width=True)
            
            col1, col2 = st.columns(2)
            if col1.button("💾 구글 시트 및 기록 저장"):
                for idx, row in edited.iterrows():
                    st.session_state.df_raw.at[idx, '리오더 수량'] = float(row['리오더 수량']) + float(row['추가 리오더'])
                save_reorder_data(st.session_state.df_raw.loc[edited.index, ['상품코드', '리오더 수량']])
                st.session_state.history[datetime.now().strftime("%Y-%m-%d %H:%M:%S")] = edited.copy()
                st.success("✅ 저장 완료!")
            
            with col2:
                csv_data = edited.to_csv(index=False).encode('utf-8-sig')
                st.download_button("📥 발주 리스트 엑셀 다운로드", csv_data, f"발주_{datetime.now().strftime('%Y%m%d')}.csv", "text/csv")

    # 6단계: 과거 데이터 확인
    st.subheader("📜 6단계: 과거 데이터 확인")
    if st.session_state.history:
        select_h = st.selectbox("⏰ 기록된 시간 선택", list(st.session_state.history.keys()))
        st.dataframe(st.session_state.history[select_h], use_container_width=True)
        csv_h = st.session_state.history[select_h].to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 선택 기록 다운로드", csv_h, f"발주기록_{select_h.replace(':', '-')}.csv", "text/csv")

