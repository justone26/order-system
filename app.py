import streamlit as st
import pandas as pd
import numpy as np
import re
import unicodedata
import gspread
from datetime import datetime, timedelta, timezone

# 1. 환경 설정
KST = timezone(timedelta(hours=9))
st.set_page_config(layout="wide", page_title="저스트원 발주 시스템")

# --- [필수 함수] ---
def find_idx(cols, keys):
    for i, c in enumerate(cols):
        if any(k in str(c) for k in keys): return i
    return 0

def super_clean(t):
    if not t: return ""
    t = unicodedata.normalize('NFC', str(t))
    return re.sub(r'[^a-zA-Z0-9가-힣]', '', t).upper().strip()

def to_i(v):
    try: return int(float(str(v).replace(",", "")))
    except: return 0

def get_sheet():
    try:
        client = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        return client.open_by_key("1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg")
    except: return None

# --- [메인 로직 시작] ---
st.header("1️⃣ 파일 업로드")
up_file = st.file_uploader("엑셀 업로드", type=['xlsx', 'xls'])

if up_file:
    if 'df_raw' not in st.session_state:
        st.session_state.df_raw = pd.read_excel(up_file)
        st.session_state.analyzed = False

    df_work = st.session_state.df_raw
    cols = df_work.columns.tolist()

    # --- [2단계: 매핑 설정 - 사장님 틀 유지] ---
    st.subheader("⚙️ 2단계: 매핑 설정")
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

    # --- [3단계: 분석 설정 - 사장님 틀 유지] ---
    st.subheader("⚙️ 3단계: 분석 설정")
    col_lt, col_ss = st.columns(2)
    lead_time = col_lt.number_input("리드타임 (일)", value=10)
    safety_stock = col_ss.number_input("안전재고 (일 수)", value=7)

    if st.button("🚀 분석 실행"):
        df = st.session_state.df_raw.copy()
        
        # 리오더 수량 실시간 반영
        sh = get_sheet()
        r_map = {}
        if sh:
            ws = sh.worksheet("발주기록")
            logs = ws.get_all_values()
            if len(logs) > 1:
                df_l = pd.DataFrame(logs[1:], columns=logs[0])
                df_l['k'] = df_l.apply(lambda r: super_clean(r[1]) + super_clean(r[2]), axis=1)
                r_map = df_l.groupby('k').apply(lambda x: x.iloc[:, 6].apply(to_i).sum()).to_dict()

        df['clean_k'] = df.apply(lambda r: super_clean(r[item]) + super_clean(r[option]), axis=1)
        df['기존리오더'] = df['clean_k'].map(r_map).fillna(0).astype(int)
        
        daily_avg = pd.to_numeric(df[t1week], errors='coerce').fillna(0) / 7
        df['권장발주수량'] = ((daily_avg * lead_time) + (daily_avg * safety_stock) - (pd.to_numeric(df[avail], errors='coerce').fillna(0) + df['기존리오더'])).clip(lower=0).astype(int)
        
        df['추가발주'] = 0
        df['입고차감'] = 0
        df['메모'] = ""
        
        st.session_state.df_raw = df
        st.session_state.analyzed = True
        st.rerun()

    # --- [4~7단계 시작] ---
    if st.session_state.get('analyzed'):
        # 4️⃣5️⃣단계: 편집 및 저장
        st.divider()
        st.subheader("4️⃣~5️⃣ 발주 수량 편집 및 저장")
        
        df_final = st.session_state.df_raw
        edited_df = st.data_editor(
            df_final,
            column_config={
                "추가발주": st.column_config.NumberColumn("➕ 추가발주"),
                "입고차감": st.column_config.NumberColumn("➖ 입고차감"),
                "메모": st.column_config.TextColumn("📝 메모")
            },
            disabled=[c for c in df_final.columns if c not in ['추가발주', '입고차감', '메모']],
            hide_index=True, use_container_width=True
        )

        if st.button("💾 일괄 저장"):
            to_save = edited_df[(edited_df['추가발주'] > 0) | (edited_df['입고차감'] > 0)]
            if not to_save.empty:
                sh = get_sheet()
                ws = sh.worksheet("발주기록")
                now = datetime.now(KST).strftime('%Y-%m-%d %H:%M')
                rows = [[
                    now, str(r[item]), str(r[option]), str(r[vendor_item]), 
                    int(to_i(r[avail])), int(r['기존리오더']), 
                    int(r['추가발주']) - int(r['입고차감']), 
                    int(r['권장발주수량']), str(r['메모']), str(r[vendor])
                ] for _, r in to_save.iterrows()]
                ws.append_rows(rows)
                st.success("✅ 저장 완료!")
                st.rerun()

        # 6️⃣단계: 저장 내역 상세 검색
        st.divider()
        st.subheader("6️⃣ 저장 내역 상세 검색")
        c6_1, c6_2 = st.columns(2)
        q_date = c6_1.date_input("날짜 선택")
        if c6_2.button("🚀 검색 실행", use_container_width=True):
            sh = get_sheet()
            if sh:
                raw_logs = sh.worksheet("발주기록").get_all_values()
                df_log = pd.DataFrame(raw_logs[1:], columns=raw_logs[0])
                target = q_date.strftime('%Y-%m-%d')
                res = df_log[df_log.iloc[:, 0].str.contains(target)].copy()
                
                # 순서 고정: 날짜, 업체명, 상품명, 옵션, 공급처상품명, 가용, 기존, 수량(G), 추가, 권장, 메모
                # (중복 제거 로직 포함)
                col_view = [res.columns[0], res.columns[9], res.columns[1], res.columns[2], res.columns[3], 
                            res.columns[4], res.columns[5], res.columns[6], res.columns[7], res.columns[8]]
                st.dataframe(res[col_view].iloc[::-1], use_container_width=True, hide_index=True)

        # 7️⃣단계: 실시간 잔량 상황판
        st.divider()
        st.subheader("7️⃣ 실시간 리오더 최종 잔량 상황판")
        if st.button("📊 현황판 업데이트"):
            sh = get_sheet()
            if sh:
                raw = sh.worksheet("발주기록").get_all_values()
                df_7 = pd.DataFrame(raw[1:], columns=raw[0])
                df_7['qty'] = df_7.iloc[:, 6].apply(to_i)
                v_sum = df_7.groupby(df_7.columns[9])['qty'].sum().reset_index()
                v_sum = v_sum[v_sum['qty'] > 0]
                
                m_cols = st.columns(4)
                for i, r in enumerate(v_sum.itertuples()):
                    with m_cols[i % 4]: st.metric(r[1], f"{int(r[2])} 개")
