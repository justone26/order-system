import streamlit as st
import pandas as pd
import numpy as np
import re
import unicodedata
import gspread
from datetime import datetime, timedelta, timezone

# 1. 환경 및 시간 설정
KST = timezone(timedelta(hours=9))
st.set_page_config(layout="wide", page_title="저스트원 발주 시스템")

# --- [공통 함수 영역: 절대 삭제 금지] ---
def super_clean(t):
    if not t: return ""
    t = unicodedata.normalize('NFC', str(t))
    return re.sub(r'[^a-zA-Z0-9가-힣]', '', t).upper().strip()

def to_i(v):
    try:
        val = str(v).replace(",", "").strip()
        return int(float(val)) if val else 0
    except: return 0

def find_idx(cols, keys):
    for i, c in enumerate(cols):
        if any(k in str(c).upper() for k in keys): return i
    return 0

def get_sheet():
    try:
        client = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        return client.open_by_key("1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg")
    except Exception as e:
        st.error(f"📡 시트 연결 실패: {e}")
        return None

# --- [메인 프로그램 시작] ---

# 1️⃣단계: 파일 업로드 및 구글 리오더 즉시 로드
st.header("1️⃣ 파일 업로드 및 리오더 데이터 로드")
up_file = st.file_uploader("엑셀 파일을 업로드하세요.", type=['xlsx', 'xls'])

if up_file:
    if 'df_raw' not in st.session_state:
        st.session_state.df_raw = pd.read_excel(up_file)
        st.session_state.analyzed = False
        
        # [핵심] 업로드 즉시 구글 리오더 값 로드
        with st.spinner("🔄 구글 시트에서 실시간 리오더 값을 가져오는 중..."):
            sh = get_sheet()
            r_map = {}
            if sh:
                try:
                    ws = sh.worksheet("발주기록")
                    logs = ws.get_all_values()
                    if len(logs) > 1:
                        df_l = pd.DataFrame(logs[1:], columns=[c.strip() for c in logs[0]])
                        df_l['k'] = df_l.apply(lambda r: super_clean(r.iloc[1]) + super_clean(r.iloc[2]), axis=1)
                        r_map = df_l.groupby('k').apply(lambda x: x.iloc[:, 6].apply(to_i).sum()).to_dict()
                except: pass
            st.session_state.r_map = r_map
        st.success("✅ 엑셀 로드 및 리오더 값 동기화 완료!")

# ------------------------------------------------------------------
# 2️⃣단계: 매핑 설정 (좌우 5개씩 총 10개 항목)
# ------------------------------------------------------------------
if 'df_raw' in st.session_state:
    st.divider()
    df_work = st.session_state.df_raw
    cols = df_work.columns.tolist()
    st.subheader("⚙️ 2️⃣단계: 매핑 설정 (10가지 항목)")
    
    c1, c2 = st.columns(2)
    with c1:
        sold_out = st.selectbox("1. 품절 여부", cols, index=find_idx(cols, ['품절']))
        vendor = st.selectbox("2. 공급처(업체명)", cols, index=find_idx(cols, ['공급처', '업체']))
        item = st.selectbox("3. 상품명", cols, index=find_idx(cols, ['상품명', '품명']))
        option = st.selectbox("4. 옵션", cols, index=find_idx(cols, ['옵션', '규격']))
        vendor_item = st.selectbox("5. 공급처 상품명", cols, index=find_idx(cols, ['공급처상품명']))
    with c2:
        reg_date = st.selectbox("6. 등록일", cols, index=find_idx(cols, ['등록일']))
        stock = st.selectbox("7. 정상재고", cols, index=find_idx(cols, ['정상재고']))
        avail = st.selectbox("8. 가용재고", cols, index=find_idx(cols, ['가용재고', '현재고']))
        t3day = st.selectbox("9. 3일 발주합계", cols, index=find_idx(cols, ['3일']))
        t1week = st.selectbox("10. 7일 발주합계", cols, index=find_idx(cols, ['7일', '1주']))

    # ------------------------------------------------------------------
    # 3️⃣단계: 분석 설정 및 실행
    # ------------------------------------------------------------------
    st.divider()
    st.subheader("⚙️ 3️⃣단계: 분석 설정")
    col_lt, col_ss = st.columns(2)
    lead_time = col_lt.number_input("리드타임 (일)", value=10)
    safety_stock = col_ss.number_input("안전재고 (일 수)", value=7)

    if st.button("🚀 분석 실행", type="primary", use_container_width=True):
        df = st.session_state.df_raw.copy()
        r_map = st.session_state.get('r_map', {})
        df['clean_k'] = df.apply(lambda r: super_clean(r[item]) + super_clean(r[option]), axis=1)
        df['기존리오더'] = df['clean_k'].map(r_map).fillna(0).astype(int)
        
        daily_avg = pd.to_numeric(df[t1week], errors='coerce').fillna(0) / 7
        df['권장발주수량'] = ((daily_avg * (lead_time + safety_stock)) - (pd.to_numeric(df[avail], errors='coerce').fillna(0) + df['기존리오더'])).clip(lower=0).astype(int)
        df['상태'] = df['권장발주수량'].apply(lambda x: "🚨 긴급" if x > 0 else "✅ 정상")
        df['추가발주'] = 0
        df['입고차감'] = 0
        df['메모'] = ""
        df = df.sort_values(by=['상태', item, option], ascending=[True, True, True])

        st.session_state.df_raw = df
        st.session_state.analyzed = True
        st.rerun()

    # ------------------------------------------------------------------
    # 4️⃣~5️⃣단계: 발주 편집 및 저장
    # ------------------------------------------------------------------
    if st.session_state.get('analyzed'):
        st.divider()
        st.header("4️⃣~5️⃣ 발주 수량 편집 및 저장")
        df_final = st.session_state.df_raw.copy()
        f1, f2, f3 = st.columns([1, 1, 2])
        sel_v = f1.selectbox("업체 필터", ["전체"] + sorted(df_final[vendor].unique().tolist()))
        sel_s = f2.selectbox("상태 필터", ["전체", "🚨 긴급", "✅ 정상"])
        q_word = f3.text_input("상품명 검색")

        disp_df = df_final.copy()
        if sel_v != "전체": disp_df = disp_df[disp_df[vendor] == sel_v]
        if sel_s != "전체": disp_df = disp_df[disp_df['상태'] == sel_s]
        if q_word: disp_df = disp_df[disp_df[item].str.contains(q_word, case=False, na=False)]

        safe_cols = ['상태', vendor, item, option, vendor_item, avail, '기존리오더', '권장발주수량', '추가발주', '입고차감', '메모']
        edited_df = st.data_editor(disp_df[safe_cols], hide_index=True, use_container_width=True, key="v16_editor")

        if st.button("💾 일괄 저장하기", type="primary", use_container_width=True):
            to_save = edited_df[(edited_df['추가발주'] > 0) | (edited_df['입고차감'] > 0)]
            if not to_save.empty:
                sh = get_sheet()
                ws = sh.worksheet("발주기록")
                now = datetime.now(KST).strftime('%Y-%m-%d %H:%M')
                rows = [[now, str(r[item]), str(r[option]), str(r[vendor_item]), int(to_i(r[avail])), int(r['기존리오더']), int(r['추가발주']) - int(r['입고차감']), int(r['권장발주수량']), str(r['메모']), str(r[vendor])] for _, r in to_save.iterrows()]
                ws.append_rows(rows)
                st.success("✅ 저장 성공!")
                st.session_state.clear()
                st.rerun()

        # ------------------------------------------------------------------
        # 6️⃣단계: 저장 내역 상세 검색
        # ------------------------------------------------------------------
        st.divider()
        st.header("6️⃣ 저장 내역 상세 검색")
        c6_1, c6_2 = st.columns(2)
        q_date = c6_1.date_input("조회 날짜 선택", key="q6_date")
        if c6_2.button("🚀 조회 실행", use_container_width=True, key="q6_btn"):
            sh = get_sheet()
            if sh:
                raw_logs = sh.worksheet("발주기록").get_all_values()
                if len(raw_logs) > 1:
                    df_log = pd.DataFrame(raw_logs[1:], columns=raw_logs[0])
                    target = q_date.strftime('%Y-%m-%d')
                    res = df_log[df_log.iloc[:, 0].str.contains(target)].copy()
                    st.dataframe(res.iloc[::-1], use_container_width=True, hide_index=True)

        # ------------------------------------------------------------------
        # 7️⃣단계: 실시간 리오더 최종 잔량 상황판
        # ------------------------------------------------------------------
        st.divider()
        st.header("7️⃣ 실시간 리오더 최종 잔량 상황판")
        if st.button("📊 현황 업데이트", use_container_width=True, key="q7_btn"):
            sh = get_sheet()
            if sh:
                raw = sh.worksheet("발주기록").get_all_values()
                if len(raw) > 1:
                    df_7 = pd.DataFrame(raw[1:], columns=raw[0])
                    df_7['qty'] = df_7.iloc[:, 6].apply(to_i)
                    v_sum = df_7.groupby(df_7.columns[-1])['qty'].sum().reset_index()
                    v_sum = v_sum[v_sum['qty'] > 0]
                    m_cols = st.columns(4)
                    for i, r in enumerate(v_sum.itertuples()):
                        with m_cols[i % 4]: st.metric(r[1], f"{int(r[2])} 개")
