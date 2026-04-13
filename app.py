import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import re
import time
import unicodedata
import gspread
from datetime import datetime, timedelta, timezone

# 1. 환경 설정
KST = timezone(timedelta(hours=9))
st.set_page_config(layout="wide", page_title="저스트원 통합 관리 v5.8")

# [새로고침 방지]
components.html("<script>window.onbeforeunload = function() { return '변경사항이 저장되지 않을 수 있습니다.'; };</script>", height=0)

# --- [사장님표 필수 함수] ---
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

# --- [메인 화면 시작] ---
st.title("📦 저스트원 통합 재고 관리 (전 과정 한 화면)")

# 1단계: 업로드
st.header("1️⃣ 데이터 업로드")
up_file = st.file_uploader("📂 엑셀 또는 CSV 파일 업로드", type=['xlsx', 'xls', 'csv'])

if up_file:
    # 데이터 로드 (최초 1회)
    if 'df_raw' not in st.session_state or st.session_state.get('last_uploaded') != up_file.name:
        with st.spinner("🚀 시트 동기화 중..."):
            df = pd.read_csv(up_file) if up_file.name.endswith('.csv') else pd.read_excel(up_file)
            df.columns = [str(c).strip() for c in df.columns]
            st.session_state.df_raw = df.fillna("")
            st.session_state.last_uploaded = up_file.name
            
            # 리오더 잔량 로드
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

    cols = st.session_state.df_raw.columns.tolist()

    # 2단계: 매핑 (사장님 요청 레이아웃)
    st.divider()
    st.header("2️⃣ 필드 매핑 설정")
    c_left, c_right = st.columns(2)
    with c_left:
        st.markdown("##### [ 기본 정보 ]")
        it = st.selectbox("📦 상품명", cols, index=auto_idx(cols, ['상품명']), key="sel_it")
        op = st.selectbox("🎨 옵션", cols, index=auto_idx(cols, ['옵션']), key="sel_op")
        vn = st.selectbox("🏭 공급처", cols, index=auto_idx(cols, ['공급처']), key="sel_vn")
        vi = st.selectbox("🆔 공급처 상품명", cols, index=auto_idx(cols, ['공급처상품명']), key="sel_vi")
        so = st.selectbox("🚫 품절 여부", cols, index=auto_idx(cols, ['품절']), key="sel_so")

    with c_right:
        st.markdown("##### [ 수량 및 날짜 ]")
        av = st.selectbox("✅ 가용재고", cols, index=auto_idx(cols, ['가용재고']), key="sel_av")
        stk = st.selectbox("📦 정상재고", cols, index=auto_idx(cols, ['정상재고']), key="sel_stk")
        
        t3_target = "3일 발주합계"
        t3_idx = cols.index(t3_target) if t3_target in cols else auto_idx(cols, ['3일'], exclude_keys=['1주', '7일', '품절'])
        t3 = st.selectbox("🔥 3일 판매", cols, index=t3_idx, key="sel_t3")
        
        t7_target = "1주발주합계"
        t7_idx = cols.index(t7_target) if t7_target in cols else auto_idx(cols, ['7일', '1주'], exclude_keys=['3일', '품절'])
        t7 = st.selectbox("📅 7일 판매", cols, index=t7_idx, key="sel_t7")
        
        reg = st.selectbox("📆 상품 등록일", cols, index=auto_idx(cols, ['등록일', '등록일자', '최초등록']), key="sel_reg")

    # 3단계: 수치 설정
    st.divider()
    st.header("3️⃣ 분석 수치 설정")
    col_lt, col_ss = st.columns(2)
    with col_lt: lt = st.number_input("⏳ 리드타임 (일)", value=7)
    with col_ss: ss = st.number_input("🛡️ 안전재고 (개)", value=3)

    # 분석 버튼 (이제 화면 전환 없이 아래에 표를 띄움)
    analyze_clicked = st.button("📊 분석 실행 (아래에 편집 창 생성)", type="primary", use_container_width=True)

    if analyze_clicked or 'analyzed_data' in st.session_state:
        if analyze_clicked:
            # 실시간 계산
            df = st.session_state.df_raw.copy()
            for c in [av, t3, t7, stk]: df[c] = df[c].apply(to_i)
            
            df['clean_key'] = df.apply(lambda r: super_clean(r[it]) + super_clean(r[op]), axis=1)
            df['기존잔량'] = df['clean_key'].map(st.session_state.r_map).fillna(0).astype(int)
            df['일판매'] = df.apply(lambda r: int(round(r[t7]/7)) if r[t7]>0 else (int(round(r[t3]/3)) if r[t3]>0 else 0), axis=1)
            df['발주권장'] = ((df['일판매'] * (lt + ss)) - (df[av] + df['기존잔량'])).clip(lower=0).astype(int)
            
            df['입고차감'] = 0
            df['추가발주'] = 0
            df['메모'] = ""
            
            st.session_state.analyzed_data = df
            st.session_state.mapping_info = {'it':it, 'op':op, 'vn':vn, 'reg':reg}

        # --- [4~7단계: 편집 및 저장 영역 (한 화면 아래에 위치)] ---
        st.divider()
        st.header("4️⃣ 발주 편집 및 💾 최종 저장")
        m = st.session_state.mapping_info
        
        edited_df = st.data_editor(
            st.session_state.analyzed_data[[m['vn'], m['it'], m['op'], '기존잔량', '입고차감', '발주권장', '추가발주', '메모', m['reg']]],
            use_container_width=True,
            hide_index=True,
            key="main_editor",
            column_config={
                "입고차감": st.column_config.NumberColumn("📥 4단계: 입고(-)", help="입고 수량"),
                "추가발주": st.column_config.NumberColumn("➕ 5단계: 추가발주", help="추가 발주량"),
                "기존잔량": st.column_config.NumberColumn("📦 현재잔량", disabled=True),
                "발주권장": st.column_config.NumberColumn("💡 권장", disabled=True)
            }
        )

        # 저장 버튼
        if st.button("💾 모든 변경사항 구글 시트 일괄 저장 (6-7단계)", type="primary", use_container_width=True):
            to_save = edited_df[(edited_df['입고차감'] != 0) | (edited_df['추가발주'] > 0)]
            if not to_save.empty:
                with st.spinner("📡 데이터 전송 중..."):
                    sh = get_sheet()
                    ws_main = sh.worksheet("발주기록")
                    ws_hist = sh.worksheet("history")
                    now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M')
                    
                    rows = []
                    for _, r in to_save.iterrows():
                        final_change = int(r['추가발주']) - int(r['입고차감'])
                        rows.append([
                            now_s, str(r[m['it']]), str(r[m['op']]), "", 0, 
                            int(r['기존잔량']), final_change, int(r['발주권장']), str(r['메모']), str(r[m['vn']])
                        ])
                    
                    ws_main.append_rows(rows)
                    ws_hist.append_rows(rows)
                    st.success("✅ 저장 완료! 리오더 수량이 갱신되었습니다.")
                    # 완료 후 세션 초기화하여 다음 작업을 준비
                    del st.session_state.analyzed_data
                    st.rerun()
            else:
                st.warning("⚠️ 입력된 데이터가 없습니다.")

        if st.button("🔄 작업 초기화 (파일 다시 올리기)"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()
