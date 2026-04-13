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

# [⚠️ 새로고침 경고] - 브라우저 종료나 새로고침 시 데이터 유실 주의 문구
st.warning("⚠️ 주의: 브라우저를 '새로고침'하면 작업 중인 데이터가 초기화됩니다. 저장을 완료한 후 페이지를 이동하세요.")

# --- [공통 함수 영역] ---
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

# 1️⃣단계: 파일 업로드 및 데이터 로드
st.header("1️⃣ 파일 업로드 및 데이터 로드")
col_up, col_reset = st.columns([4, 1])

with col_up:
    up_file = st.file_uploader("엑셀 파일을 업로드하세요.", type=['xlsx', 'xls'])

with col_reset:
    if st.button("🔄 전체 화면 초기화", help="모든 설정을 리셋하고 처음으로 돌아갑니다."):
        st.session_state.clear()
        st.rerun()

if up_file:
    if 'df_raw' not in st.session_state:
        st.session_state.df_raw = pd.read_excel(up_file)
        st.session_state.analyzed = False
        with st.spinner("🔄 구글 시트 리오더 데이터 실시간 동기화 중..."):
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
        st.success("✅ 파일 로드 및 리오더 값 동기화 완료!")

# ------------------------------------------------------------------
# 2️⃣단계: 매핑 설정 (5:5 배치)
# ------------------------------------------------------------------
if 'df_raw' in st.session_state:
    st.divider()
    df_work = st.session_state.df_raw
    cols = df_work.columns.tolist()
    st.subheader("⚙️ 2️⃣단계: 매핑 설정")
    
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
    # 3️⃣단계: 분석 설정 및 실행 (등록일 기준 판매량 보정 로직 포함)
    # ------------------------------------------------------------------
    st.divider()
    st.subheader("⚙️ 3️⃣단계: 분석 설정")
    col_lt, col_ss = st.columns(2)
    lead_time = col_lt.number_input("리드타임 (일)", value=10)
    safety_stock = col_ss.number_input("안전재고 (일 수)", value=7)

    if st.button("🚀 분석 실행", type="primary", use_container_width=True):
        df = st.session_state.df_raw.copy()
        r_map = st.session_state.get('r_map', {})
        today = datetime.now(KST).date()

        # [핵심 로직] 등록일 기준 일평균 판매량 계산
        def calc_daily_avg(row):
            try:
                # 등록일 파싱
                r_dt = pd.to_datetime(row[reg_date]).date()
                diff_days = (today - r_dt).days
                if diff_days < 1: diff_days = 1 # 오늘 등록했으면 1일로 계산
                
                sum_7d = to_i(row[t1week])
                
                # 등록한지 7일이 안 되었으면 등록일 기준으로 일평균 계산
                actual_days = min(diff_days, 7)
                return sum_7d / actual_days
            except:
                return to_i(row[t1week]) / 7

        df['일평균판매량'] = df.apply(calc_daily_avg, axis=1)
        df['clean_k'] = df.apply(lambda r: super_clean(r[item]) + super_clean(r[option]), axis=1)
        df['기존리오더'] = df['clean_k'].map(r_map).fillna(0).astype(int)
        
        # 권장 발주량 계산
        avail_val = pd.to_numeric(df[avail], errors='coerce').fillna(0)
        df['권장발주수량'] = ((df['일평균판매량'] * (lead_time + safety_stock)) - (avail_val + df['기존리오더'])).clip(lower=0).astype(int)
        
        # 상태값 정의
        def get_status(row):
            if "품절" in str(row[sold_out]): return "🚫 품절"
            return "🚨 긴급" if row['권장발주수량'] > 0 else "✅ 정상"
        df['상태'] = df.apply(get_status, axis=1)
        
        # [긴급 상품 묶음 정렬 로직]
        is_emergency_item = df.groupby(item)['상태'].transform(lambda x: any(x == "🚨 긴급"))
        df['sort_group'] = np.where(is_emergency_item, 0, 1) # 긴급 포함 상품이 0번 그룹
        
        df['추가발주'] = 0
        df['입고차감'] = 0
        df['메모'] = ""
        
        # 최종 정렬
        df = df.sort_values(by=['sort_group', item, option], ascending=[True, True, True])

        st.session_state.df_raw = df
        st.session_state.analyzed = True
        st.rerun()

    # ------------------------------------------------------------------
    # 4️⃣단계: 발주 편집 및 저장 (통합)
    # ------------------------------------------------------------------
    if st.session_state.get('analyzed'):
        st.divider()
        st.header("4️⃣ 발주 수량 편집 및 저장")
        
        df_final = st.session_state.df_raw.copy()
        
        # 필터링
        f1, f2 = st.columns([1, 2])
        sel_s = f1.selectbox("🚦 상태 필터", ["전체상품", "🚨 긴급", "✅ 정상", "🚫 품절"])
        q_word = f2.text_input("🔎 상품명 검색")

        disp_df = df_final.copy()
        if sel_s == "🚨 긴급": disp_df = disp_df[disp_df['sort_group'] == 0]
        elif sel_s == "✅ 정상": disp_df = disp_df[disp_df['상태'] == "✅ 정상"]
        elif sel_s == "🚫 품절": disp_df = disp_df[disp_df['상태'] == "🚫 품절"]
        
        if q_word:
            disp_df = disp_df[disp_df[item].str.contains(q_word, case=False, na=False)]

        safe_cols = ['상태', vendor, item, option, vendor_item, avail, '기존리오더', '권장발주수량', '추가발주', '입고차감', '메모']
        
        st.info("💡 Tip: 긴급 상품은 모든 옵션이 세트로 노출됩니다. 추가발주/입고차감 입력 후 저장하세요.")
        edited_df = st.data_editor(
            disp_df[safe_cols],
            column_config={
                "상태": st.column_config.TextColumn("상태", width="small"),
                item: st.column_config.TextColumn("상품명", width="large"),
                "추가발주": st.column_config.NumberColumn("➕ 추가발주", step=1),
                "입고차감": st.column_config.NumberColumn("➖ 입고차감", step=1),
            },
            disabled=[c for c in safe_cols if c not in ['추가발주', '입고차감', '메모']],
            hide_index=True, use_container_width=True, key="v19_final_editor"
        )

        if st.button("💾 위 내역 일괄 저장", type="primary", use_container_width=True):
            to_save = edited_df[(edited_df['추가발주'] > 0) | (edited_df['입고차감'] > 0)]
            if not to_save.empty:
                sh = get_sheet()
                ws = sh.worksheet("발주기록")
                now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M')
                rows = [[now_s, str(r[item]), str(r[option]), str(r[vendor_item]), int(to_i(r[avail])), int(r['기존리오더']), int(r['추가발주']) - int(r['입고차감']), int(r['권장발주수량']), str(r['메모']), str(r[vendor])] for _, r in to_save.iterrows()]
                ws.append_rows(rows)
                st.success("✅ 구글 시트 저장 완료! 화면을 새로고침합니다.")
                st.session_state.clear()
                st.rerun()

        # 6-7단계 (검색 및 현황판)
        st.divider()
        c6, c7 = st.columns(2)
        with c6:
            st.subheader("6️⃣ 내역 검색")
            q_date = st.date_input("날짜 선택", key="search_date")
            if st.button("검색 실행", key="search_btn"):
                sh = get_sheet()
                if sh:
                    logs = pd.DataFrame(sh.worksheet("발주기록").get_all_values())
                    logs.columns = logs.iloc[0]; logs = logs[1:]
                    res = logs[logs.iloc[:, 0].str.contains(q_date.strftime('%Y-%m-%d'))]
                    st.dataframe(res.iloc[::-1], use_container_width=True, hide_index=True)
        with c7:
            st.subheader("7️⃣ 잔량 현황")
            if st.button("현황 업데이트", key="status_btn"):
                sh = get_sheet(); raw = pd.DataFrame(sh.worksheet("발주기록").get_all_values())
                raw.columns = raw.iloc[0]; raw = raw[1:]
                raw['q'] = raw.iloc[:, 6].apply(to_i)
                v_sum = raw.groupby(raw.columns[-1])['q'].sum().reset_index()
                for r in v_sum[v_sum['q']>0].itertuples(): st.write(f"**{r[1]}**: {int(r[2])}개")
