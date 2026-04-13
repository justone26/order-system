import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import re
import unicodedata
import gspread
from datetime import datetime, timedelta, timezone

# 1. 환경 설정
KST = timezone(timedelta(hours=9))
st.set_page_config(layout="wide", page_title="저스트원 v6.9")

# [새로고침 방지]
components.html("<script>window.onbeforeunload = function() { return '변경사항이 저장되지 않을 수 있습니다.'; };</script>", height=0)

# --- [공통 함수] ---
def super_clean(t):
    if not t: return ""
    t = unicodedata.normalize('NFC', str(t))
    return re.sub(r'[^a-zA-Z0-9가-힣]', '', t).upper().strip()

def to_i(v):
    try: return int(float(str(v).replace(",", "").strip()))
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
        c_str = str(c).upper()
        if exclude_keys and any(k.upper() in c_str for k in exclude_keys): continue
        if any(k.upper() in c_str for k in keys): return i
    return 0

# --- [메인 화면] ---
st.title("📦 저스트원 통합 재고 관리")

# 1단계: 업로드 & 2단계: 매핑 & 3단계: 분석수치 (기존 코드와 동일하므로 분석 로직부터 집중 업데이트)
# ... (중간 생략: 1~3단계 기본 UI) ...

if 'df_raw' in st.session_state:
    cols = st.session_state.df_raw.columns.tolist()
    st.divider()
    st.header("2️⃣ 필드 매핑 & 3️⃣ 수치 설정")
    c1, c2, c3 = st.columns([2, 2, 1])
    with c1:
        sel_it = st.selectbox("📦 상품명", cols, index=auto_idx(cols, ['상품명']), key="it_box")
        sel_op = st.selectbox("🎨 옵션", cols, index=auto_idx(cols, ['옵션']), key="op_box")
        sel_vn = st.selectbox("🏭 공급처", cols, index=auto_idx(cols, ['공급처']), key="vn_box")
        sel_vi = st.selectbox("🆔 공급처 상품명", cols, index=auto_idx(cols, ['공급처상품명']), key="vi_box")
    with c2:
        sel_av = st.selectbox("✅ 가용재고", cols, index=auto_idx(cols, ['가용재고']), key="av_box")
        sel_so = st.selectbox("🚫 품절", cols, index=auto_idx(cols, ['품절']), key="so_box")
        sel_t3 = st.selectbox("🔥 3일 판매", cols, index=auto_idx(cols, ['3일']), key="t3_box")
        sel_t7 = st.selectbox("📅 7일 판매", cols, index=auto_idx(cols, ['7일']), key="t7_box")
    with c3:
        lt = st.number_input("⏳ 리드타임", value=7)
        ss = st.number_input("🛡️ 안전재고", value=3)

    if st.button("📊 분석 실행", type="primary", use_container_width=True):
        df = st.session_state.df_raw.copy()
        for c in [sel_av, sel_t3, sel_t7]: df[c] = df[c].apply(to_i)
        
        # 기본 계산
        df['clean_key'] = df.apply(lambda r: super_clean(r[sel_it]) + super_clean(r[sel_op]), axis=1)
        df['리오더 수량'] = df['clean_key'].map(st.session_state.r_map).fillna(0).astype(int)
        df['1일 판매량'] = df.apply(lambda r: int(round(r[sel_t7]/7)) if r[sel_t7]>0 else (int(round(r[sel_t3]/3)) if r[sel_t3]>0 else 0), axis=1)
        df['권장발주수량'] = ((df['1일 판매량'] * (lt + ss)) - (df[sel_av] + df['리오더 수량'])).clip(lower=0).astype(int)
        
        # 긴급발주 여부 판단 (권장발주가 0보다 크면 긴급)
        df['is_urgent'] = df['권장발주수량'] > 0
        
        # [핵심] 긴급발주인 상품의 '전체 옵션'을 묶기 위한 작업
        urgent_items = df[df['is_urgent'] == True][sel_it].unique()
        df['item_urgent_group'] = df[sel_it].isin(urgent_items)
        
        df['입고차감'] = 0
        df['추가발주'] = 0
        df['메모'] = ""
        
        # 정렬: 긴급그룹 우선 -> 상품명 -> 옵션 순
        df = df.sort_values(by=['item_urgent_group', sel_it, sel_op], ascending=[False, True, True])
        
        st.session_state.final_mapping = {'vn': sel_vn, 'it': sel_it, 'op': sel_op, 'vi': sel_vi, 'av': sel_av, 't3': sel_t3, 'so': sel_so}
        st.session_state.analyzed_data = df

# --- [4~5단계: 편집 및 필터 영역] ---
if 'analyzed_data' in st.session_state:
    st.divider()
    st.header("4️⃣~5️⃣ 발주 편집 및 필터")
    
    # [상단 필터바]
    f1, f2 = st.columns([1, 1])
    with f1:
        filter_type = st.radio("🔍 보기 필터", ["전체 보기", "긴급발주(묶음)", "정상재고", "품절"], horizontal=True)
    with f2:
        search_it = st.text_input("🔎 상품명 검색", placeholder="검색할 상품명을 입력하세요...")

    # 필터링 로직
    m = st.session_state.final_mapping
    df_f = st.session_state.analyzed_data.copy()

    # 1. 상태 필터
    if filter_type == "긴급발주(묶음)":
        df_f = df_f[df_f['item_urgent_group'] == True]
    elif filter_type == "정상재고":
        df_f = df_f[df_f[m['so']].get(str, "").str.contains("정상", na=False)]
    elif filter_type == "품절":
        df_f = df_f[df_f[m['so']].get(str, "").str.contains("품절", na=False)]

    # 2. 검색 필터
    if search_it:
        df_f = df_f[df_f[m['it']].str.contains(search_it, case=False, na=False)]

    # 표 구성
    display_cols = [m['vn'], m['it'], m['op'], m['vi'], m['av'], '리오더 수량', '입고차감', '추가발주', m['t3'], '1일 판매량', '권장발주수량', '메모']
    
    # 긴급건 색상 강조를 위해 에디터 설정
    edited_df = st.data_editor(
        df_f[display_cols],
        use_container_width=True, hide_index=True, key="main_editor",
        column_config={
            "권장발주수량": st.column_config.NumberColumn("💡 권장", help="0보다 크면 긴급"),
            "입고차감": st.column_config.NumberColumn("📥 입고(-)")
        }
    )

    if st.button("💾 변경사항 일괄 저장", type="primary", use_container_width=True):
        # 저장 로직 (기존과 동일)
        to_save = edited_df[(edited_df['입고차감'] != 0) | (edited_df['추가발주'] > 0)]
        if not to_save.empty:
            # ... (저장 코드 실행) ...
            st.success("✅ 저장 완료!")
            del st.session_state.analyzed_data
            st.rerun()

        # --- [6단계: 히스토리 (최근 저장 내역)] ---
        st.divider()
        st.header("6️⃣ 최근 히스토리 (History)")
        sh_hist = get_sheet()
        if sh_hist:
            ws_h = sh_hist.worksheet("history")
            # 최근 10개 행만 가져오기
            hist_data = ws_h.get_all_values()
            if len(hist_data) > 1:
                h_df = pd.DataFrame(hist_data[1:], columns=hist_data[0]).tail(10)
                st.table(h_df.iloc[::-1]) # 최신순 정렬해서 표로 표시
            else:
                st.write("기록된 히스토리가 없습니다.")

        # --- [7단계: 리오더 현황판 (실시간)] ---
        st.divider()
        st.header("7️⃣ 실시간 리오더 현황판")
        if 'r_map' in st.session_state:
            # 잔량이 있는 항목만 모아서 요약
            summary = []
            for k, v in st.session_state.r_map.items():
                if v > 0: summary.append({"상품 식별키": k, "현재 리오더 총 잔량": v})
            
            if summary:
                st.dataframe(pd.DataFrame(summary), use_container_width=True)
            else:
                st.info("현재 진행 중인 리오더 잔량이 없습니다.")
