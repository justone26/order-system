import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import re
import time
import unicodedata
import gspread
from datetime import datetime, timedelta, timezone

# 1. 환경 설정 및 시간 정의
KST = timezone(timedelta(hours=9))
st.set_page_config(layout="wide", page_title="저스트원 v5.5")

# [새로고침 방지 스크립트] - 최상단 배치
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

# --- [세션 상태 관리] ---
if 'step' not in st.session_state: st.session_state.step = 1
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'r_map' not in st.session_state: st.session_state.r_map = {}
if 'mapping' not in st.session_state: st.session_state.mapping = {}

st.title("📦 저스트원 통합 재고 관리")

# --- [1단계: 엑셀 업로드] ---
if st.session_state.step == 1:
    st.header("1️⃣ 엑셀 업로드 및 실시간 데이터 동기화")
    up_file = st.file_uploader("파일을 선택하세요 (xlsx, csv)", type=['xlsx', 'xls', 'csv'])
    
    if up_file:
        with st.spinner("🚀 데이터를 읽고 구글 시트 리오더 잔량을 가져오는 중..."):
            df = pd.read_csv(up_file) if up_file.name.endswith('.csv') else pd.read_excel(up_file)
            df.columns = [str(c).strip() for c in df.columns]
            st.session_state.df_raw = df.fillna("")
            
            # 구글 시트 리오더 데이터 1회 로드 (4-7단계 딜레이 해결 핵심)
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
            
            st.success("✅ 파일 로드 및 시트 동기화 완료!")
            if st.button("2단계: 매핑 설정으로 이동 ➡️"):
                st.session_state.step = 2
                st.rerun()

# --- [2단계: 필수 매핑 10개 항목 - 좌우 5:5] ---
elif st.session_state.step == 2:
    st.header("2️⃣ 필드 매핑 설정 (10개 필수 항목)")
    cols = st.session_state.df_raw.columns.tolist()
    
    m1, m2 = st.columns(2)
    with m1:
        it = st.selectbox("1. 📦 상품명", cols, index=0)
        op = st.selectbox("2. 🎨 옵션", cols, index=1)
        vn = st.selectbox("3. 🏭 공급처", cols, index=2)
        vi = st.selectbox("4. 🆔 공급처상품명", cols, index=3)
        rg = st.selectbox("5. 📝 등록이", cols, index=0)
    with m2:
        av = st.selectbox("6. ✅ 가용재고", cols, index=4)
        t3 = st.selectbox("7. 🔥 3일 판매량", cols, index=5)
        t7 = st.selectbox("8. 📅 7일 판매량", cols, index=6)
        pr = st.selectbox("9. 💰 판매가", cols, index=0)
        ct = st.selectbox("10. 📁 카테고리/기타", cols, index=0)

    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        if st.button("⬅️ 이전 단계 (업로드)"):
            st.session_state.step = 1
            st.rerun()
    with col_btn2:
        if st.button("3단계: 수치 설정으로 이동 ➡️"):
            st.session_state.mapping = {
                'it':it, 'op':op, 'vn':vn, 'vi':vi, 'rg':rg,
                'av':av, 't3':t3, 't7':t7, 'pr':pr, 'ct':ct
            }
            st.session_state.step = 3
            st.rerun()

# --- [3단계: 수치 설정 및 분석 실행] ---
elif st.session_state.step == 3:
    st.header("3️⃣ 분석 수치 설정 및 계산")
    c1, c2 = st.columns(2)
    with c1:
        lt = st.number_input("⏳ 리드타임 (일)", value=7)
    with c2:
        ss = st.number_input("🛡️ 안전재고 (개)", value=3)
    
    st.divider()
    if st.button("📊 분석 시작 (결과 화면으로)", type="primary", use_container_width=True):
        m = st.session_state.mapping
        df = st.session_state.df_raw.copy()
        
        # 숫자 정제
        for c in [m['av'], m['t3'], m['t7']]:
            df[c] = df[c].apply(to_i)
        
        # 리오더 잔량 계산
        df['clean_key'] = df.apply(lambda r: super_clean(r[m['it']]) + super_clean(r[m['op']]), axis=1)
        df['기존잔량'] = df['clean_key'].map(st.session_state.r_map).fillna(0).astype(int)
        
        # 사장님 권장발주식 적용
        df['일판매'] = df.apply(lambda r: int(round(r[m['t7']]/7)) if r[m['t7']]>0 else (int(round(r[m['t3']]/3)) if r[m['t3']]>0 else 0), axis=1)
        df['발주권장'] = ((df['일판매'] * (lt + ss)) - (df[m['av']] + df['기존잔량'])).clip(lower=0).astype(int)
        
        df['입고차감'] = 0
        df['추가발주'] = 0
        df['메모'] = ""
        
        st.session_state.df_final = df
        st.session_state.step = 4
        st.rerun()

# --- [4~7단계: 편집 및 최종 일괄 저장] ---
elif st.session_state.step == 4:
    st.header("4️⃣ 발주 편집 및 저장 (4~7단계)")
    m = st.session_state.mapping
    
    # 데이터 에디터 (수정 중 딜레이 없음)
    edited_df = st.data_editor(
        st.session_state.df_final[[m['vn'], m['it'], m['op'], '기존잔량', '입고차감', '발주권장', '추가발주', '메모', m['rg']]],
        use_container_width=True,
        hide_index=True,
        column_config={
            "입고차감": st.column_config.NumberColumn("📥 4단계: 입고(-)", help="입고 수량"),
            "추가발주": st.column_config.NumberColumn("➕ 5단계: 추가발주", help="신규 발주 수량"),
            "기존잔량": st.column_config.NumberColumn("📦 현재잔량", disabled=True)
        }
    )

    col_save, col_reset = st.columns(2)
    with col_save:
        if st.button("💾 구글 시트 일괄 저장 (6~7단계)", type="primary", use_container_width=True):
            to_save = edited_df[(edited_df['입고차감'] != 0) | (edited_df['추가발주'] > 0)]
            if not to_save.empty:
                with st.spinner("📡 구글 시트에 기록 중..."):
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
                    
                    # 일괄 전송으로 딜레이 최소화
                    ws_main.append_rows(rows)
                    ws_hist.append_rows(rows)
                    st.success("✅ 저장이 완료되었습니다. 리오더 수량이 갱신되었습니다!")
                    time.sleep(1)
                    st.session_state.step = 1
                    st.rerun()
            else:
                st.warning("⚠️ 저장할 변경 데이터가 없습니다.")
    
    with col_reset:
        if st.button("🔄 전체 초기화", use_container_width=True):
            st.session_state.step = 1
            st.session_state.df_raw = None
            st.rerun()
