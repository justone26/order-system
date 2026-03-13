import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json

# --- 여기부터 함수 정의 시작 ---
def get_sheet():
    creds_dict = dict(st.secrets["gcp_service_account"])
    scope = ["https://spreadsheets.google.com/feeds", 
             'https://www.googleapis.com/auth/spreadsheets', 
             "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    spreadsheet_key = "1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg"
    return client.open_by_key(spreadsheet_key).sheet1

def load_reorder_data():
    try:
        sheet = get_sheet()
        return pd.DataFrame(sheet.get_all_records())
    except: 
        return pd.DataFrame(columns=['상품코드', '1차 리오더', '2차 리오더'])

def save_reorder_data(df):
    sheet = get_sheet()
    sheet.clear()
    sheet.update([df.columns.values.tolist()] + df.values.tolist())

def get_auto_index(cols, keywords):
    for key in keywords:
        for i, c in enumerate(cols):
            if key in str(c): return i
    return 0
# --- 함수 정의 끝 ---

# 이제 여기서부터 앱의 메인 로직이 시작되어야 해!
st.set_page_config(layout="wide", page_title="재고 관리 시스템")
# ... 나머지 코드들 ...

# [상태 초기화]
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'history' not in st.session_state: st.session_state.history = {}

st.title("📦 재고 관리 및 발주 시스템")

if st.button("🔄 시스템 전체 초기화"):
    for key in list(st.session_state.keys()): del st.session_state[key]
    st.rerun()

# [파일 업로드]
uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'xls', 'csv'])
if uploaded_file is not None and st.session_state.df_raw is None:
    df = pd.read_excel(uploaded_file)
    reorder_df = load_reorder_data()
    if not reorder_df.empty:
        df = df.merge(reorder_df[['상품코드', '1차 리오더', '2차 리오더']], on='상품코드', how='left').fillna(0)
    else:
        df['1차 리오더'] = 0; df['2차 리오더'] = 0
    st.session_state.df_raw = df.loc[:, ~df.columns.duplicated()]
    st.rerun()

if st.session_state.df_raw is not None:
    cols = st.session_state.df_raw.columns.tolist()
    
 # 1단계: 매핑 설정
    st.subheader("⚙️ 1단계: 자동 매핑 설정")
    c1, c2 = st.columns(2)
    sold_out = c1.selectbox("품절 여부", cols, index=get_auto_index(cols, ['품절', '판매중단']))
    vendor = c1.selectbox("공급처", cols, index=get_auto_index(cols, ['공급처', '업체명']))
    item = c1.selectbox("상품명", cols, index=get_auto_index(cols, ['상품명', '상품']))
    option = c1.selectbox("옵션", cols, index=get_auto_index(cols, ['옵션']))
    vendor_item = c1.selectbox("공급처 상품명", cols, index=get_auto_index(cols, ['공급처상품명', '거래처옵션']))
    
    reg_date = c2.selectbox("등록일", cols, index=get_auto_index(cols, ['등록일', '생성일']))
    stock = c2.selectbox("정상재고", cols, index=get_auto_index(cols, ['정상재고', '재고']))
    avail = c2.selectbox("가용재고", cols, index=get_auto_index(cols, ['가용재고', '가용']))
    t3day = c2.selectbox("3일 발주합계", cols, index=get_auto_index(cols, ['3일']))
    t1week = c2.selectbox("7일 발주합계", cols, index=get_auto_index(cols, ['7일', '1주']))

# 2~3단계: 분석 설정
    st.subheader("⚙️ 2~3단계: 분석 설정")
    col1, col2 = st.columns(2)
    
    # 초기값: 리드타임 10, 안전재고 7
    lead_time = col1.number_input("리드타임 (일)", min_value=0, value=10)
    safety_stock = col2.number_input("안전재고 (일 수)", min_value=0, value=7)

    if st.button("🚀 분석 실행"):
        df = st.session_state.df_raw.copy()
        
        # 7일 발주합계를 기준으로 일일 판매량 계산
        daily_avg = pd.to_numeric(df[t1week], errors='coerce').fillna(0) / 7
        
        # 계산 공식: (일일판매량 * 리드타임) + (일일판매량 * 안전재고일수) - 가용재고
        needed = (daily_avg * lead_time) + (daily_avg * safety_stock)
        current_inv = pd.to_numeric(df[avail], errors='coerce').fillna(0)
        
        df['권장 발주량'] = (needed - current_inv).clip(lower=0).astype(int)
        st.session_state.df_raw = df
        st.success("분석 완료!")
        st.rerun()

# 4단계: 데이터 편집 및 리오더 관리
    st.subheader("📊 4단계: 데이터 편집 및 리오더 관리")

    # [필수 열 생성/확인]
    for col in ['리오더 수량', '입고 완료', '권장 발주량']:
        if col not in st.session_state.df_raw.columns:
            st.session_state.df_raw[col] = 0

    # --- 검색 및 필터 설정 구역 ---
    col_search, col_filter = st.columns([2, 1])
    search_term = col_search.text_input("🔍 상품명 또는 옵션 검색", "")

    status_option = col_filter.selectbox(
        "판매 상태 필터",
        ["정상 (빈칸)", "품절 (글자있음)", "전체보기"],
        index=2  # 데이터가 안 보일 때는 일단 '전체보기'로 index를 2로 바꿔서 테스트해봐!
    )

    # --- 데이터 필터링 로직 ---
    display_df = st.session_state.df_raw.copy()

    # 1. 판매 상태 필터링 (NaN 및 공백 처리 강화)
    if status_option == "정상 (빈칸)":
        # 문자열로 변환 후 앞뒤 공백 제거했을 때 빈 값인 것만 추출
        display_df = display_df[
            (display_df[sold_out].isna()) | 
            (display_df[sold_out].astype(str).str.strip() == "") |
            (display_df[sold_out].astype(str).str.strip() == "nan")
        ]
    elif status_option == "품절 (글자있음)":
        # 문자열이 있고, 'nan'이 아니며, 공백이 아닌 것
        display_df = display_df[
            (display_df[sold_out].notna()) & 
            (display_df[sold_out].astype(str).str.strip() != "") &
            (display_df[sold_out].astype(str).str.strip() != "nan")
        ]

    # 2. 검색어 필터링
    if search_term:
        display_df = display_df[
            display_df[item].astype(str).str.contains(search_term, case=False, na=False) | 
            display_df[option].astype(str).str.contains(search_term, case=False, na=False)
        ]

    # --- 데이터 편집기 출력 ---
    display_cols = [
        sold_out, vendor, item, option, vendor_item, 
        stock, avail, t1week, 
        '리오더 수량', '입고 완료', '권장 발주량'
    ]

    if not display_df.empty:
        edited_df = st.data_editor(
            display_df[display_cols], 
            use_container_width=True,
            key="data_editor_main"
        )
        if edited_df is not None:
            # 수정된 내용 반영 (index 기준으로 업데이트)
            st.session_state.df_raw.update(edited_df)
    else:
        st.info("💡 해당 조건의 데이터가 없습니다. 필터를 '전체보기'로 변경해보세요.")
        # 만약 전체보기에서도 안 나오면 매핑 문제!
        if status_option == "전체보기" and st.session_state.df_raw.empty:
             st.error("⚠️ 원본 데이터가 비어있습니다. 1단계 파일을 다시 확인해주세요.")
                
    # 5단계: 발주 리스트 요약 (에러 방어 버전)
    st.subheader("📋 5단계: 발주 리스트 요약")
    
    # [방어 로직] 데이터프레임에 '권장 발주량' 컬럼이 존재하는지 먼저 확인
    if '권장 발주량' in st.session_state.df_raw.columns:
        to_order = st.session_state.df_raw[st.session_state.df_raw['권장 발주량'] > 0].copy()
        
        if not to_order.empty:
            st.warning(f"🚨 발주 대상 {len(to_order)}개 확인")
            st.dataframe(to_order[[vendor, item, option, "1차 리오더", "2차 리오더", "권장 발주량"]], use_container_width=True)
            
            if st.button("💾 구글 시트 및 기록 저장"):
                save_reorder_data(to_order[['상품코드', '1차 리오더', '2차 리오더']])
                st.session_state.history[datetime.now().strftime("%Y-%m-%d %H:%M:%S")] = to_order.copy()
                st.success("저장 완료!")
        else:
            st.info("✅ 현재 발주할 상품이 없습니다.")
    else:
        st.info("💡 '분석 실행' 버튼을 누르면 발주 리스트가 나타납니다.")

    # [6단계: 과거 기록]
    st.subheader("📜 6단계: 과거 데이터 확인")
    if st.session_state.history:
        select_h = st.selectbox("⏰ 시간 선택", list(st.session_state.history.keys()))
        st.dataframe(st.session_state.history[select_h])
















