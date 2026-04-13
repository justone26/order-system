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
        # 사장님 시트 키값 확인 필요
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

# ------------------------------------------------------------------
# 1️⃣단계: 파일 업로드
# ------------------------------------------------------------------
st.header("1️⃣ 엑셀 파일 업로드")
up_file = st.file_uploader("재고/판매 엑셀 파일을 업로드하세요.", type=['xlsx', 'xls'])

if up_file:
    df_raw = pd.read_excel(up_file)
    st.session_state.df_raw = df_raw
    st.success("✅ 파일 로드 완료!")

# ------------------------------------------------------------------
# 2️⃣단계: 컬럼 설정 & 3️⃣단계: 분석 실행 (같은 화면 선상)
# ------------------------------------------------------------------
if 'df_raw' in st.session_state:
    st.divider()
    cols = st.session_state.df_raw.columns.tolist()
    
    c_setup, c_action = st.columns([3, 1])
    
    with c_setup:
        st.subheader("2️⃣ 분석 컬럼 설정")
        row1_1, row1_2, row1_3 = st.columns(3)
        with row1_1:
            s_vn = st.selectbox("업체명", cols, index=auto_idx(cols, ['공급처', '업체']), key="v2_vn")
            s_it = st.selectbox("상품명", cols, index=auto_idx(cols, ['상품명', '품명']), key="v2_it")
        with row1_2:
            s_op = st.selectbox("옵션", cols, index=auto_idx(cols, ['옵션', '규격']), key="v2_op")
            s_vi = st.selectbox("공급처상품명", cols, index=auto_idx(cols, ['공급처상품명']), key="v2_vi")
        with row1_3:
            s_av = st.selectbox("가용재고", cols, index=auto_idx(cols, ['가용재고', '현재고']), key="v2_av")
            s_t7 = st.selectbox("7일판매량", cols, index=auto_idx(cols, ['7일', '판매']), key="v2_t7")
            s_t3 = st.selectbox("3일판매량", cols, index=auto_idx(cols, ['3일']), key="v2_t3")

    with c_action:
        st.subheader("3️⃣ 분석 실행")
        lt = st.number_input("리드타임", value=7, key="v3_lt")
        ss = st.number_input("안전재고", value=3, key="v3_ss")
        if st.button("📊 분석 실행", type="primary", use_container_width=True):
            df = st.session_state.df_raw.copy()
            
            # 리오더 수량 로드 (r_map)
            sh = get_sheet()
            r_map = {}
            if sh:
                ws_log = sh.worksheet("발주기록")
                all_logs = ws_log.get_all_values()
                if len(all_logs) > 1:
                    df_log = pd.DataFrame(all_logs[1:], columns=all_logs[0])
                    df_log['key'] = df_log.apply(lambda r: super_clean(r[1]) + super_clean(r[2]), axis=1)
                    r_map = df_log.groupby('key').apply(lambda x: x.iloc[:, 6].apply(to_i).sum()).to_dict()

            # 데이터 가공
            df['clean_key'] = df.apply(lambda r: super_clean(r[s_it]) + super_clean(r[s_op]), axis=1)
            df['리오더 수량'] = df['clean_key'].map(r_map).fillna(0).astype(int)
            
            # 권장 수량 및 긴급 필터
            df['1일 판매량'] = df.apply(lambda r: int(round(to_i(r[s_t7])/7)) if to_i(r[s_t7])>0 else int(round(to_i(r[s_t3])/3)), axis=1)
            df['권장발주수량'] = ((df['1일 판매량'] * (lt + ss)) - (df[s_av].apply(to_i) + df['리오더 수량'])).clip(lower=0).astype(int)
            df['상태'] = df['권장발주수량'].apply(lambda x: "🚨 긴급" if x > 0 else "✅ 정상")
            
            # 긴급 그룹 우선 정렬
            urgent_items = df[df['상태'] == "🚨 긴급"][s_it].unique()
            df['is_urgent'] = df[s_it].isin(urgent_items)
            df = df.sort_values(['is_urgent', s_it], ascending=[False, True])
            
            df['추가발주'] = 0
            df['입고차감'] = 0
            df['메모'] = ""
            
            st.session_state.analyzed_data = df
            st.session_state.final_mapping = {'vn':s_vn, 'it':s_it, 'op':s_op, 'vi':s_vi, 'av':s_av, 'vn_name':s_vn}
            st.rerun()

# ------------------------------------------------------------------
# 4️⃣~5️⃣단계: 발주 편집 및 저장 (독립 구역)
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
        hide_index=True, use_container_width=True, key="main_editor"
    )

    if st.button("💾 일괄 저장 (구글 시트)", type="primary", use_container_width=True):
        to_save = edited_df[(edited_df['입고차감'] > 0) | (edited_df['추가발주'] > 0)]
        if not to_save.empty:
            sh = get_sheet()
            ws = sh.worksheet("발주기록")
            now = datetime.now(KST).strftime('%Y-%m-%d %H:%M')
            rows = [[
                now, str(r[m['it']]), str(r[m['op']]), str(r[m['vi']]), 
                int(to_i(r[m['av']])), int(r['리오더 수량']), 
                int(r['추가발주']) - int(r['입고차감']), 
                int(r['권장발주수량']), str(r['메모']), str(r[m['vn']])
            ] for _, r in to_save.iterrows()]
            ws.append_rows(rows)
            st.success("✅ 저장 완료! 수치 갱신 중...")
            if 'r_map' in st.session_state: del st.session_state.r_map
            st.rerun()

# ------------------------------------------------------------------
# 6️⃣단계: 저장 내역 상세 검색 (독립 구역)
# ------------------------------------------------------------------
if 'analyzed_data' in st.session_state:
    st.divider()
    st.header("6️⃣ 저장 내역 상세 검색")
    
    c6_1, c6_2 = st.columns([1, 1])
    with c6_1:
        q_date_6 = st.date_input("📅 날짜 선택", key="q_date_6")
    with c6_2:
        st.write("")
        btn_6 = st.button("🚀 데이터 조회하기", use_container_width=True, type="primary", key="btn_6")

    if btn_6 or st.session_state.get('load6'):
        st.session_state.load6 = True
        sh = get_sheet()
        if sh:
            ws = sh.worksheet("발주기록")
            raw = ws.get_all_values()
            if len(raw) > 1:
                df6 = pd.DataFrame(raw[1:], columns=[c.strip() for c in raw[0]])
                df6['p_date'] = df6.iloc[:, 0].str.split(' ').str[0]
                df6['p_time'] = df6.iloc[:, 0].str.split(' ').str[1].str[:5]
                
                target = q_date_6.strftime('%Y-%m-%d')
                res = df6[df6['p_date'] == target].copy()
                
                if not res.empty:
                    f1, f2 = st.columns(2)
                    with f1:
                        ts = sorted(res['p_time'].unique(), reverse=True)
                        q_t = st.selectbox(f"⏰ 회차 선택 ({len(ts)}회)", ["전체"] + ts)
                    with f2:
                        q_i = st.text_input("🔎 상품명 검색")
                    
                    if q_t != "전체": res = res[res['p_time'] == q_t]
                    if q_i: res = res[res.iloc[:, 1].str.contains(q_i, case=False)]
                    
                    # 순서 고정: 날짜, 업체명(끝), 상품명(1), 옵션(2), 공급처상품(3), 가용(4), 기존(5), 입고/추가(6), 권장(7), 메모(8)
                    cols_final = [res.columns[0], res.columns[-3], res.columns[1], res.columns[2], res.columns[3], 
                                  res.columns[4], res.columns[5], res.columns[6], "추가발주", "권장발주수량", "메모"]
                    # 중복 제거 및 존재하는 컬럼만
                    show_cols = []
                    for c in cols_final:
                        if c in res.columns and c not in show_cols: show_cols.append(c)
                    
                    st.dataframe(res[show_cols].iloc[::-1], use_container_width=True, hide_index=True)
                else: st.info("데이터가 없습니다.")

# ------------------------------------------------------------------
# 7️⃣단계: 실시간 리오더 최종 잔량 상황판 (독립 구역)
# ------------------------------------------------------------------
if 'analyzed_data' in st.session_state:
    st.divider()
    st.header("7️⃣ 실시간 리오더 최종 잔량 상황판")
    
    if st.button("📊 실시간 현황판 업데이트", use_container_width=True, key="btn_7"):
        sh = get_sheet()
        if sh:
            raw = sh.worksheet("발주기록").get_all_values()
            if len(raw) > 1:
                df7 = pd.DataFrame(raw[1:], columns=raw[0])
                df7['qty'] = df7.iloc[:, 6].apply(to_i)
                v_col = df7.columns[-1]
                
                df_rem = df7.groupby([v_col, df7.columns[1], df7.columns[2]], as_index=False)['qty'].sum()
                df_rem = df_rem[df_rem['qty'] > 0]
                
                if not df_rem.empty:
                    v_sum = df_rem.groupby(v_col)['qty'].sum().reset_index()
                    m_cols = st.columns(4)
                    for idx, r in enumerate(v_sum.itertuples()):
                        with m_cols[idx % 4]: st.metric(r[1], f"{int(r[2])} 개")
                    st.dataframe(df_rem, use_container_width=True, hide_index=True)
                else: st.success("✅ 미입고 잔량 없음!")
