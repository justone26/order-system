import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import re
import unicodedata
import gspread
from datetime import datetime, timedelta, timezone
import io

# 1. 환경 설정
KST = timezone(timedelta(hours=9))
st.set_page_config(layout="wide", page_title="저스트원 v9.2")

# [새로고침 방지]
components.html("<script>window.onbeforeunload = function() { return '변경사항이 저장되지 않을 수 있습니다.'; };</script>", height=0)

# --- [공통 함수] ---
def super_clean(t):
    if not t: return ""
    t = unicodedata.normalize('NFC', str(t))
    return re.sub(r'[^a-zA-Z0-9가-힣]', '', t).upper().strip()

def to_i(v):
    try: 
        val = str(v).replace(",", "").strip()
        return int(float(val)) if val else 0
    except: return 0

def get_sheet():
    try:
        client = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        return client.open_by_key("1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg")
    except Exception as e:
        st.error(f"📡 시트 연결 실패: {e}")
        return None

def auto_idx(cols, keys, exclude_keys=None):
    for i, c in enumerate(cols):
        c_str = str(c).upper().replace(" ", "")
        if exclude_keys and any(k.upper() in c_str for k in exclude_keys): continue
        if any(k.upper() in c_str for k in keys): return i
    return 0

# --- [공통 데이터 로드] ---
# 여기서부터는 왼쪽 벽에 딱 붙여야 합니다! (공백 제거)
st.divider()
sh = get_sheet()

if sh:
    ws_log = sh.worksheet("발주기록")
    raw_logs = ws_log.get_all_values()
    
    # 데이터가 있는 경우에만 아래 단계들 실행
    if len(raw_logs) > 1:
        df_logs = pd.DataFrame(raw_logs[1:], columns=[c.strip() for c in raw_logs[0]])
        d_col = next((c for c in df_logs.columns if '날짜' in c), df_logs.columns[0])
        v_col = next((c for c in df_logs.columns if '공급처' in c), None)

        # 날짜/시간 전처리 (공통)
        df_logs['pure_dt'] = df_logs[d_col].str.strip()
        df_logs['pure_date'] = df_logs['pure_dt'].str.split(' ').str[0]
        df_logs['pure_time'] = df_logs['pure_dt'].str.split(' ').str[1].str[:5]


# --- [메인 화면] ---
st.title("📦 저스트원 통합 재고 관리")

# ------------------------------------------------------------------
# 1️⃣단계: 파일 업로드
# ------------------------------------------------------------------
st.header("1️⃣ 엑셀 파일 업로드")
up_file = st.file_uploader("재고/판매 엑셀 파일을 업로드하세요.", type=['xlsx', 'xls'])

if up_file:
    df_raw = pd.read_excel(up_file)
    st.session_state.df_raw = df_raw
    st.success("✅ 파일 업로드 완료!")

# ------------------------------------------------------------------
# 2️⃣단계: 컬럼 설정 & 3️⃣단계: 분석 실행 (같은 화면 선상 배치)
# ------------------------------------------------------------------
if 'df_raw' in st.session_state:
    st.divider()
    cols = st.session_state.df_raw.columns.tolist()
    
    col_setup, col_action = st.columns([3, 1])
    
    with col_setup:
        st.subheader("2️⃣ 분석 컬럼 설정")
        c1, c2, c3 = st.columns(3)
        with c1:
            s_vn = st.selectbox("업체명", cols, index=auto_idx(cols, ['공급처', '업체']))
            s_it = st.selectbox("상품명", cols, index=auto_idx(cols, ['상품명', '품명']))
        with c2:
            s_op = st.selectbox("옵션", cols, index=auto_idx(cols, ['옵션', '규격']))
            s_vi = st.selectbox("공급처상품명", cols, index=auto_idx(cols, ['공급처상품명']))
        with c3:
            s_av = st.selectbox("가용재고", cols, index=auto_idx(cols, ['가용재고', '현재고']))
            s_t7 = st.selectbox("7일판매량", cols, index=auto_idx(cols, ['7일', '판매']))
            
    with col_action:
        st.subheader("3️⃣ 분석 실행")
        lt = st.number_input("리드타임", value=7)
        ss = st.number_input("안전재고", value=3)
        btn_analyze = st.button("📊 분석 실행", type="primary", use_container_width=True)

    if btn_analyze:
        # [기존 분석 로직 실행 및 긴급 우선 정렬 코드 포함 구역]
        # (이 부분은 사장님이 가지고 계신 긴급 정렬 로직을 그대로 사용하시면 됩니다.)
        # ... 분석 후 st.session_state.analyzed_data 저장 ...
        pass
        

# ------------------------------------------------------------------
# 4️⃣~5️⃣단계: 발주 편집 및 일괄 저장 (독립 블록)
# ------------------------------------------------------------------
if 'analyzed_data' in st.session_state:
    st.divider()
    st.header("4️⃣~5️⃣ 발주 수량 편집 및 저장")
    
    df_res = st.session_state.analyzed_data
    m = st.session_state.final_mapping

    edited_df = st.data_editor(
        df_res,
        column_config={
            "추가발주": st.column_config.NumberColumn("➕ 추가발주", min_value=0),
            "입고차감": st.column_config.NumberColumn("➖ 입고차감", min_value=0),
            "메모": st.column_config.TextColumn("📝 메모")
        },
        disabled=[c for c in df_res.columns if c not in ['추가발주', '입고차감', '메모']],
        hide_index=True, use_container_width=True, key="editor_v10"
    )

    if st.button("💾 일괄 저장 (구글 시트 전송)", type="primary", use_container_width=True):
        # ... 저장 및 st.rerun() 로직 ...
        pass

        
  # ------------------------------------------------------------------
# 6️⃣단계: 저장 내역 상세 검색 (독립 블록)
# ------------------------------------------------------------------
if 'analyzed_data' in st.session_state:
    st.divider()
    st.header("6️⃣ 저장 내역 상세 검색")
    
    c6_1, c6_2 = st.columns([1, 1])
    with c6_1:
        q_date_6 = st.date_input("📅 날짜 선택", key="search_date_6")
    with c6_2:
        st.write("")
        btn_load_6 = st.button("🚀 내역 조회하기", use_container_width=True, key="btn_6")

    if btn_load_6 or st.session_state.get('load6'):
        st.session_state.load6 = True
        # ... 시트 읽기 및 날짜/회차/상품명 필터 로직 ...
        # ... display_6[final_cols] 출력 ...
        pass

        

# ------------------------------------------------------------------
# 7️⃣단계: 실시간 리오더 최종 잔량 상황판 (독립 블록)
# ------------------------------------------------------------------
if 'analyzed_data' in st.session_state:
    st.divider()
    st.header("7️⃣ 실시간 리오더 최종 잔량 상황판")
    
    if st.button("📊 실시간 현황판 업데이트", use_container_width=True, key="btn_7"):
        # ... 시트 전체 읽기 및 미입고 잔량 집계 로직 ...
        # ... st.metric() 및 현황 표 출력 ...
        pass
