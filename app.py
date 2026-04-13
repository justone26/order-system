import streamlit as st
import pandas as pd
import numpy as np
import re
import unicodedata
import gspread
from datetime import datetime, timedelta, timezone
import streamlit.components.v1 as components

# 1. 환경 및 시간 설정
KST = timezone(timedelta(hours=9))
st.set_page_config(layout="wide", page_title="저스트원 발주 시스템")

# --- [🚨 새로고침/창닫기 방지 자바스크립트] ---
components.html(
    """
    <script>
    window.addEventListener('beforeunload', function (e) {
      e.preventDefault();
      e.returnValue = '';
    });
    </script>
    """,
    height=0,
)

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

def load_reorder_data():
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
    return r_map

# --- [메인 로직] ---

# 1️⃣단계: 파일 업로드
st.header("1️⃣ 파일 업로드 및 데이터 로드")
up_file = st.file_uploader("엑셀 파일을 업로드하세요.", type=['xlsx', 'xls'])

if st.button("🔄 현재 화면 데이터 초기화", use_container_width=True):
    st.session_state.clear()
    st.rerun()

if up_file:
    if 'df_raw' not in st.session_state:
        st.session_state.df_raw = pd.read_excel(up_file)
        st.session_state.r_map = load_reorder_data()
        st.session_state.analyzed = False
        st.success("✅ 파일 로드 및 리오더 값 동기화 완료!")

# 2️⃣단계: 매핑 설정
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
        v_item_col = st.selectbox("5. 공급처 상품명", cols, index=find_idx(cols, ['공급처상품명']))
    with c2:
        reg_date = st.selectbox("6. 등록일", cols, index=find_idx(cols, ['등록일']))
        stock = st.selectbox("7. 정상재고", cols, index=find_idx(cols, ['정상재고']))
        avail = st.selectbox("8. 가용재고", cols, index=find_idx(cols, ['가용재고', '현재고']))
        t3d = st.selectbox("9. 3일 발주합계", cols, index=find_idx(cols, ['3일']))
        t1w = st.selectbox("10. 7일 발주합계", cols, index=find_idx(cols, ['7일', '1주']))

    # 3️⃣단계: 분석 설정
    st.divider()
    st.subheader("⚙️ 3️⃣단계: 분석 설정")
    clt, css = st.columns(2)
    lt, ss = clt.number_input("리드타임 (일)", value=10), css.number_input("안전재고 (일 수)", value=7)

    if st.button("🚀 분석 실행 / 업데이트", type="primary", use_container_width=True):
        df = st.session_state.df_raw.copy()
        r_map = load_reorder_data()
        today = datetime.now(KST).date()

        def get_daily_avg(row):
            try:
                r_dt = pd.to_datetime(row[reg_date]).date()
                diff = (today - r_dt).days
                days = max(1, min(diff, 7)) 
                return int(round(to_i(row[t1w]) / days))
            except: 
                return int(round(to_i(row[t1w]) / 7))

        df['일판매량'] = df.apply(get_daily_avg, axis=1)
        df['clean_k'] = df.apply(lambda r: super_clean(r[item]) + super_clean(r[option]), axis=1)
        df['기존리오더'] = df['clean_k'].map(r_map).fillna(0).astype(int)
        
        av_val = pd.to_numeric(df[avail], errors='coerce').fillna(0)
        df['권장발주수량'] = ((df['일판매량'] * (lt + ss)) - (av_val + df['기존리오더'])).clip(lower=0).astype(int)
        
        def status_check(row):
            if "품절" in str(row[sold_out]): return "🚫 품절"
            return "🚨 긴급" if row['권장발주수량'] > 0 else "✅ 정상"
        df['상태'] = df.apply(status_check, axis=1)
        
        is_emg = df.groupby(item)['상태'].transform(lambda x: any(x == "🚨 긴급"))
        df['sort_group'] = np.where(is_emg, 0, 1)
        df = df.sort_values(by=['sort_group', item, option], ascending=[True, True, True])
        
        df['추가발주'], df['입고차감'], df['메모'] = 0, 0, ""
        st.session_state.df_final = df
        st.session_state.analyzed = True
        st.rerun()

    # 4️⃣, 6️⃣, 7️⃣단계: 분석 완료 시 항상 표시
    if st.session_state.get('analyzed'):
        # --- 4️⃣ 발주 편집 및 저장 ---
        st.divider()
        st.header("4️⃣ 발주 편집 및 저장")
        df_f = st.session_state.df_final.copy()
        
        f1, f2 = st.columns([1, 2])
        sel_s = f1.selectbox("🚦 상태 필터", ["전체상품", "🚨 긴급", "✅ 정상", "🚫 품절"])
        q = f2.text_input("🔎 상품명 검색")

        disp = df_f.copy()
        if sel_s == "🚨 긴급": disp = disp[disp['sort_group'] == 0]
        elif sel_s == "✅ 정상": disp = disp[disp['상태'] == "✅ 정상"]
        elif sel_s == "🚫 품절": disp = disp[disp['상태'] == "🚫 품절"]
        if q: disp = disp[disp[item].str.contains(q, case=False, na=False)]

        safe = ['상태', vendor, item, option, v_item_col, avail, '기존리오더', '입고차감', '추가발주', t3d, '일판매량', '권장발주수량', '메모']
        
        try:
            edit_df = st.data_editor(
                disp[safe], 
                hide_index=True, 
                use_container_width=True, 
                key="final_v28_full",
                column_config={
                    "상태": st.column_config.TextColumn("상태", width="small"),
                    vendor: st.column_config.TextColumn("공급처"),
                    item: st.column_config.TextColumn("상품명", width="large"),
                    v_item_col: st.column_config.TextColumn("공급처상품명"),
                    "입고차감": st.column_config.NumberColumn("➖ 입고수량", step=1),
                    "추가발주": st.column_config.NumberColumn("➕ 추가발주", step=1),
                    t3d: st.column_config.NumberColumn("📅 3일판매"),
                    "일판매량": st.column_config.NumberColumn("🔥 일판매량", format="%d"),
                    "메모": st.column_config.TextColumn("비고", width="medium")
                }
            )

            if st.button("💾 내역 저장 및 수치 즉시 갱신", type="primary", use_container_width=True):
                to_save = edit_df[(edit_df['추가발주'] > 0) | (edit_df['입고차감'] > 0)]
                if not to_save.empty:
                    sh = get_sheet()
                    ws = sh.worksheet("발주기록")
                    now = datetime.now(KST).strftime('%Y-%m-%d %H:%M')
                    # 저장 순서 고정: 발주시간, 상품명, 옵션, 공급처상품명, 가용재고, 리오더잔량, 추가발주, 발주권장, 메모, 업체명
                    rows = [[now, str(r[item]), str(r[option]), str(r[v_item_col]), int(to_i(r[avail])), int(r['기존리오더']), int(r['추가발주']) - int(r['입고차감']), int(r['권장발주수량']), str(r['메모']), str(r[vendor])] for _, r in to_save.iterrows()]
                    ws.append_rows(rows)
                    st.success("✅ 저장 완료! 리오더 수치를 업데이트합니다.")
                    st.session_state.analyzed = False 
                    st.cache_data.clear() # 캐시 초기화하여 7단계 갱신 유도
                    st.rerun()
        except KeyError as e:
            st.error(f"❌ 매핑 에러: {e} 컬럼을 찾을 수 없습니다.")

        # --- 6️⃣단계: 전체 히스토리 관리 ---
        st.divider()
        st.header("📜 6단계: 전체 히스토리 관리")
        
        h1, h2, h3, h4 = st.columns([1.2, 0.8, 1.5, 1.5])
        with h1:
            today = datetime.now(KST).date()
            d_range = st.date_input("🗓️ 1. 조회 범위", value=(today, today), key="v6_dr")
        with h2:
            st.write(""); st.write("")
            search_trigger = st.button("🔍 2. 내역 조회", use_container_width=True, type="primary")

        if 'v6_data' not in st.session_state: st.session_state.v6_data = None
        if 'v6_sessions' not in st.session_state: st.session_state.v6_sessions = []

        if search_trigger:
            sh = get_sheet()
            if sh:
                all_h = sh.worksheet("발주기록").get_all_values()
                if len(all_h) > 1:
                    df_all = pd.DataFrame(all_h[1:], columns=["발주시간", "상품명", "옵션", "공급처상품명", "가용재고", "리오더잔량", "추가발주", "발주권장", "메모", "업체명"])
                    df_all["날짜_만"] = df_all["발주시간"].str.slice(0, 10)
                    s_d = d_range[0].strftime('%Y-%m-%d'); e_d = d_range[1].strftime('%Y-%m-%d') if len(d_range)>1 else s_d
                    df_filt = df_all[(df_all["날짜_만"] >= s_d) & (df_all["날짜_만"] <= e_d)].copy()
                    st.session_state.v6_data = df_filt
                    st.session_state.v6_sessions = sorted(df_filt["발주시간"].unique(), reverse=True)
                else: st.info("저장된 내역이 없습니다.")

        with h3: h_q = st.text_input("🔍 3. 상품명/옵션 검색", key="v6_q")
        with h4:
            if st.session_state.v6_sessions:
                s_opts = ["📊 선택 범위 전체 합산"] + [f"{len(st.session_state.v6_sessions)-i}회차 ({t[5:16]})" for i, t in enumerate(st.session_state.v6_sessions)]
                sel_label = st.selectbox("📦 4. 회차 선택", s_opts)
            else: st.selectbox("📦 4. 회차 선택", ["조회 결과 없음"], disabled=True); sel_label = None

        if st.session_state.v6_data is not None and sel_label:
            df_disp = st.session_state.v6_data.copy()
            for c in ["가용재고", "리오더잔량", "추가발주", "발주권장"]: df_disp[c] = pd.to_numeric(df_disp[c], errors='coerce').fillna(0)
            
            if sel_label == "📊 선택 범위 전체 합산":
                df_disp = df_disp.groupby(["업체명", "상품명", "옵션", "공급처상품명"], as_index=False).agg({"발주시간":"max", "가용재고":"last", "리오더잔량":"last", "추가발주":"sum", "발주권장":"last", "메모": lambda x: " / ".join(set(filter(None, x.astype(str))))})
            else:
                t_time = st.session_state.v6_sessions[s_opts.index(sel_label)-1]
                df_disp = df_disp[df_disp["발주시간"] == t_time].copy()

            if h_q: df_disp = df_disp[df_disp["상품명"].str.contains(h_q, case=False) | df_disp["옵션"].str.contains(h_q, case=False)]
            st.dataframe(df_disp[["발주시간", "업체명", "상품명", "옵션", "공급처상품명", "가용재고", "리오더잔량", "추가발주", "발주권장", "메모"]], use_container_width=True, hide_index=True)

        # --- 7️⃣단계: 실시간 리오더 최종 잔량 상황판 ---
        st.divider()
        st.header("🚀 7단계: 실시간 리오더 최종 잔량 상황판")
        
        @st.cache_data(ttl=600)
        def get_v7():
            try:
                data = get_sheet().worksheet("발주기록").get_all_values()
                if len(data) > 1:
                    df = pd.DataFrame(data[1:], columns=["발주시간", "상품명", "옵션", "공급처상품명", "가용재고", "기존리오더", "추가발주", "발주권장", "메모", "업체명"])
                    df["기존리오더"] = pd.to_numeric(df["기존리오더"], errors='coerce').fillna(0).astype(int)
                    df["추가발주"] = pd.to_numeric(df["추가발주"], errors='coerce').fillna(0).astype(int)
                    df["최종잔량"] = df["기존리오더"] + df["추가발주"]
                    df["날짜_순수"] = df["발주시간"].str.slice(0, 10)
                    return df
            except: return None

        df_v7 = get_v7()
        if df_v7 is not None:
            f1, f2, f3, f4 = st.columns([1.2, 0.6, 1.5, 1.2])
            with f1: d7 = st.date_input("🗓️ 기간", value=((datetime.now(KST)-timedelta(days=30)).date(), today), key="v7_dr")
            with f2:
                st.write(""); st.write("")
                if st.button("🔄 업데이트", key="v7_up"): st.cache_data.clear(); st.rerun()
            with f3: q7 = st.text_input("🔍 검색", key="v7_qs")
            with f4: v7 = st.selectbox("🏭 업체", ["전체 업체"] + sorted(df_v7["업체명"].unique().tolist()))

            df_f7 = df_v7.copy()
            if isinstance(d7, (list, tuple)) and len(d7) == 2:
                df_f7 = df_f7[(df_f7["날짜_순수"] >= d7[0].strftime('%Y-%m-%d')) & (df_f7["날짜_순수"] <= d7[1].strftime('%Y-%m-%d'))]
            if v7 != "전체 업체": df_f7 = df_f7[df_f7["업체명"] == v7]
            if q7: df_f7 = df_f7[df_f7["상품명"].str.contains(q7, case=False) | df_f7["옵션"].str.contains(q7, case=False)]

            if not df_f7.empty:
                df_res7 = df_f7.groupby(["업체명", "상품명", "옵션", "공급처상품명"], as_index=False).agg({"발주시간":"max", "최종잔량":"sum", "메모": lambda x: " / ".join(dict.fromkeys(filter(None, x.astype(str))))})
                df_res7["최종잔량"] = df_res7["최종잔량"].clip(lower=0)
                
                df_v_sum = df_res7.groupby("업체명")["최종잔량"].sum().reset_index()
                df_v_sum = df_v_sum[df_v_sum["최종잔량"] > 0].sort_values("최종잔량", ascending=False)
                
                v_cols = st.columns(4)
                for i, r in enumerate(df_v_sum.itertuples()):
                    with v_cols[i % 4]: st.metric(r.업체명, f"{int(r.최종잔량):,} 개")
                
                st.write("#### 📋 상세 리스트 (미입고)")
                st.dataframe(df_res7[df_res7["최종잔량"] > 0].sort_values("발주시간", ascending=False), use_container_width=True, hide_index=True)
