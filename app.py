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
st.set_page_config(layout="wide", page_title="저스트원 v7.7")

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

# 1단계: 업로드
st.header("1️⃣ 데이터 업로드")
if 'reset_trigger' not in st.session_state: st.session_state.reset_trigger = 0
up_file = st.file_uploader("📂 파일 업로드", type=['xlsx', 'xls', 'csv'], key=f"uploader_{st.session_state.reset_trigger}")

if up_file:
    if 'df_raw' not in st.session_state:
        try:
            df = pd.read_csv(up_file) if up_file.name.endswith('.csv') else pd.read_excel(up_file)
            df.columns = [str(c).strip() for c in df.columns]
            st.session_state.df_raw = df.fillna("")
            sh = get_sheet()
            if sh:
                ws = sh.worksheet("발주기록")
                all_vals = ws.get_all_values()
                r_map = {}
                if len(all_vals) > 1:
                    for row in all_vals[1:]:
                        key = super_clean(row[1]) + super_clean(row[2])
                        r_map[key] = r_map.get(key, 0) + (to_i(row[5]) + to_i(row[6]))
                st.session_state.r_map = r_map
        except Exception as e:
            st.error(f"⚠️ 파일 로드 실패: {e}")
            st.stop()

    cols = st.session_state.df_raw.columns.tolist()
    
    # 2~3단계: 설정 영역 (상단 고정)
    st.divider()
    st.header("2️⃣ 필드 매핑 및 3️⃣ 수치 설정")
    c1, c2, c3 = st.columns([2, 2, 1])
    with c1:
        s_it = st.selectbox("📦 상품명", cols, index=auto_idx(cols, ['상품명']), key="it_box")
        s_op = st.selectbox("🎨 옵션", cols, index=auto_idx(cols, ['옵션']), key="op_box")
        s_vn = st.selectbox("🏭 공급처", cols, index=auto_idx(cols, ['공급처']), key="vn_box")
        s_vi = st.selectbox("🆔 공급처 상품명", cols, index=auto_idx(cols, ['공급처상품명']), key="vi_box")
    with c2:
        s_av = st.selectbox("✅ 가용재고", cols, index=auto_idx(cols, ['가용재고']), key="av_box")
        s_so = st.selectbox("🚫 품절여부", cols, index=auto_idx(cols, ['품절']), key="so_box")
        s_t3 = st.selectbox("🔥 3일 판매", cols, index=auto_idx(cols, ['3일']), key="t3_box")
        s_t7 = st.selectbox("📅 7일 판매", cols, index=auto_idx(cols, ['7일', '1주']), key="t7_box")
    with c3:
        lt = st.number_input("⏳ 리드타임", value=7, key="lt_val")
        ss = st.number_input("🛡️ 안전재고", value=3, key="ss_val")

    if st.button("📊 분석 실행", type="primary", use_container_width=True):
        df = st.session_state.df_raw.copy()
        for c in [s_av, s_t3, s_t7]: df[c] = df[c].apply(to_i)
        
        df['clean_key'] = df.apply(lambda r: super_clean(r[s_it]) + super_clean(r[s_op]), axis=1)
        df['리오더 수량'] = df['clean_key'].map(st.session_state.r_map).fillna(0).astype(int)
        df['1일 판매량'] = df.apply(lambda r: int(round(r[s_t7]/7)) if r[s_t7]>0 else (int(round(r[s_t3]/3)) if r[s_t3]>0 else 0), axis=1)
        df['권장발주수량'] = ((df['1일 판매량'] * (lt + ss)) - (df[s_av] + df['리오더 수량'])).clip(lower=0).astype(int)
        
        df['상태'] = df['권장발주수량'].apply(lambda x: "🚨 긴급" if x > 0 else "✅ 정상")
        urgent_items = df[df['권장발주수량'] > 0][s_it].unique()
        df['item_urgent_group'] = df[s_it].isin(urgent_items)
        
        df['입고차감'] = 0
        df['추가발주'] = 0
        df['메모'] = ""
        
        st.session_state.analyzed_data = df.sort_values(by=['item_urgent_group', s_it, s_op], ascending=[False, True, True])
        st.session_state.final_mapping = {'vn':s_vn, 'it':s_it, 'op':s_op, 'vi':s_vi, 'av':s_av, 't3':s_t3, 'so':s_so}

    # 4~5단계: 결과 영역 (들여쓰기 교정 완료)
    if 'analyzed_data' in st.session_state:
        st.divider()
        st.header("4️⃣~5️⃣ 발주 편집 및 저장")
        
        m = st.session_state.final_mapping
        
        f1, f2 = st.columns([3, 2])
        with f1:
            f_type = st.radio("🔍 필터", ["전체", "🚨 긴급(묶음)", "✅ 정상", "🚫 품절"], horizontal=True, key="f_radio")
        with f2:
            search_q = st.text_input("🔎 상품명 검색", key="s_input")

        df_view = st.session_state.analyzed_data.copy()
        if f_type == "🚨 긴급(묶음)":
            df_view = df_view[df_view['item_urgent_group'] == True]
        elif f_type == "✅ 정상":
            df_view = df_view[df_view[m['so']].astype(str).str.contains("정상", na=False)]
        elif f_type == "🚫 품절":
            df_view = df_view[df_view[m['so']].astype(str).str.contains("품절", na=False)]
        
        if search_q:
            df_view = df_view[df_view[m['it']].str.contains(search_q, case=False, na=False)]

        t_cols = ['상태', m['vn'], m['it'], m['op'], m['vi'], m['av'], '리오더 수량', '입고차감', '추가발주', m['t3'], '1일 판매량', '권장발주수량', '메모']
        a_cols = [c for c in t_cols if c in df_view.columns]

        edited_df = st.data_editor(
            df_view[a_cols],
            use_container_width=True, hide_index=True, key="main_editor",
            column_config={
                "상태": st.column_config.TextColumn("📊 상태", width="small"),
                "입고차감": st.column_config.NumberColumn("📥 입고(-)", format="%d"),
                "추가발주": st.column_config.NumberColumn("➕ 추가", format="%d"),
            }
        )

        if st.button("💾 일괄 저장", type="primary", use_container_width=True):
            to_save = edited_df[(edited_df['입고차감'] != 0) | (edited_df['추가발주'] > 0)]
            if not to_save.empty:
                sh = get_sheet()
                ws_main = sh.worksheet("발주기록")
                now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M')
                rows = [[now_s, str(r[m['it']]), str(r[m['op']]), str(r[m['vi']]), 0, int(r['리오더 수량']), int(r['추가발주'])-int(r['입고차감']), int(r['권장발주수량']), str(r['메모']), str(r[m['vn']])] for _, r in to_save.iterrows()]
                ws_main.append_rows(rows)
                st.success("✅ 저장 성공!")
                del st.session_state.analyzed_data
                st.rerun()

# [사이드바 초기화 버튼]
if st.sidebar.button("🔄 전체 초기화"):
    for k in list(st.session_state.keys()): del st.session_state[k]
    st.rerun()
