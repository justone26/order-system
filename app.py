import streamlit as st
import pandas as pd
import numpy as np
import time
import io
import re
import unicodedata
import streamlit.components.v1 as components
import gspread
from datetime import datetime, timedelta, timezone

# --- [1. 기본 설정 및 시간] ---
KST = timezone(timedelta(hours=9)) 
current_today = datetime.now(KST).date()
st.set_page_config(layout="wide", page_title="저스트원 재고관리 v4.0")

# 새로고침 방지 스크립트
components.html("<script>window.onbeforeunload = function() { return '변경사항이 저장되지 않을 수 있습니다.'; };</script>", height=0)

# --- [2. 세션 상태 초기화] ---
for key in ['df_raw', 'analyzed', 'p', 'add_order_dict', 'upload_key', 'v6_data']:
    if key not in st.session_state:
        if key == 'analyzed': st.session_state[key] = False
        elif key == 'p' or key == 'add_order_dict': st.session_state[key] = {}
        elif key == 'upload_key': st.session_state[key] = 0
        else: st.session_state[key] = None

# --- [3. 구글 시트 연동 최적화 (핵심)] ---

@st.cache_resource
def get_gspread_client():
    """구글 클라이언트 연결 유지 (인증 부하 방지)"""
    return gspread.service_account_from_dict(st.secrets["gcp_service_account"])

def get_sheet():
    """시트 객체 반환"""
    try:
        client = get_gspread_client()
        return client.open_by_key("1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg")
    except Exception as e:
        st.error(f"📡 시트 연결 실패: {e}")
        return None

@st.cache_data(ttl=600) # 10분간 데이터 유지
def get_cached_values(sheet_name):
    """시트 데이터를 캐싱하여 API 호출 최소화"""
    sh = get_sheet()
    if sh:
        try:
            ws = sh.worksheet(sheet_name)
            return ws.get_all_values()
        except: return []
    return []

# --- [4. 유틸리티 함수] ---

def super_clean(t):
    """상품명/옵션 텍스트 정규화"""
    if not t: return ""
    t = unicodedata.normalize('NFC', str(t))
    return re.sub(r'[^a-zA-Z0-9가-힣]', '', t).upper().strip()

def to_int(v):
    """숫자 변환 안전 장치"""
    try: return int(float(str(v).replace(",", "").strip()))
    except: return 0

# --- [5. 비즈니스 로직 함수] ---

def sync_reorder_from_sheet(df_uploaded):
    """1단계: 업로드 파일과 시트 리오더 수량 매칭"""
    all_data = get_cached_values("발주기록")
    if len(all_data) <= 1: return df_uploaded

    header = [str(h).strip().replace(" ", "") for h in all_data[0]]
    idx_name = next((i for i, h in enumerate(header) if "상품명" in h), 1)
    idx_opt = next((i for i, h in enumerate(header) if "옵션" in h), 2)
    idx_f = next((i for i, h in enumerate(header) if "기존" in h), 5)
    idx_g = next((i for i, h in enumerate(header) if "추가" in h), 6)

    reorder_map = {}
    for row in all_data[1:]:
        s_key = super_clean(row[idx_name]) + "_" + super_clean(row[idx_opt])
        qty = to_int(row[idx_f]) + to_int(row[idx_g])
        if qty != 0:
            reorder_map[s_key] = reorder_map.get(s_key, 0) + qty

    def final_match(r):
        u_key = super_clean(r['상품명']) + "_" + super_clean(r['옵션'])
        return reorder_map.get(u_key, 0)

    df_uploaded['리오더 수량'] = df_uploaded.apply(final_match, axis=1)
    return df_uploaded

def get_realtime_data_v4(target_date):
    """4단계/7단계용 실시간 잔량 계산"""
    d7 = get_cached_values("발주기록")
    dh = get_cached_values("입고기록")
    
    r_map, h_map = {}, {}
    # 발주 합산
    if len(d7) > 1:
        for row in d7[1:]:
            key = super_clean(row[1]) + super_clean(row[2])
            val = to_int(row[5]) + to_int(row[6])
            r_map[key] = r_map.get(key, 0) + val
            
    # 입고 합산 (날짜 필터)
    t_str = target_date.strftime('%Y-%m-%d')
    if len(dh) > 1:
        for row in dh[1:]:
            if t_str in str(row[0]):
                h_key = super_clean(row[1]) + super_clean(row[2])
                h_map[h_key] = h_map.get(h_key, 0) + to_int(row[3])
    return r_map, h_map

# --- [6. 메인 UI 및 단계별 로직] ---

st.title("📦 저스트원 통합 재고 관리 v4.0")
tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 관리"])

with tab1:
    # --- 1단계 ---
    st.subheader("📁 1단계: 데이터 업로드")
    up_file = st.file_uploader("엑셀/CSV 파일 업로드", type=['xlsx', 'xls', 'csv'], key=f"up_{st.session_state.upload_key}")
    
    if st.button("🔄 화면 전체 초기화", use_container_width=True):
        st.session_state.upload_key += 1
        st.session_state.df_raw = None
        st.session_state.analyzed = False
        st.cache_data.clear()
        st.rerun()

    if up_file and st.session_state.df_raw is None:
        with st.spinner("🔄 데이터 로드 및 시트 매칭 중..."):
            df = pd.read_csv(up_file) if up_file.name.endswith('.csv') else pd.read_excel(up_file)
            df.columns = [str(c).strip() for c in df.columns]
            df = sync_reorder_from_sheet(df)
            st.session_state.df_raw = df.fillna("")

    # --- 2~3단계 ---
    if st.session_state.df_raw is not None:
        st.divider()
        st.subheader("📋 2단계: 매핑 및 분석")
        cols = st.session_state.df_raw.columns.tolist()
        
        def auto_idx(keys, exclude=None):
            for i, c in enumerate(cols):
                if exclude and any(e in str(c) for e in exclude): continue
                if any(k in str(c) for k in keys): return i
            return 0

        cl, cr = st.columns(2)
        with cl:
            it = st.selectbox("📦 상품명", cols, index=auto_idx(['상품명']), key="sel_it")
            op = st.selectbox("🎨 옵션", cols, index=auto_idx(['옵션']), key="sel_op")
            vn = st.selectbox("🏭 공급처", cols, index=auto_idx(['공급처']), key="sel_vn")
            vi = st.selectbox("🆔 공급처 상품명", cols, index=auto_idx(['공급처상품명']), key="sel_vi")
            so = st.selectbox("🚫 품절 여부", cols, index=auto_idx(['품절']), key="sel_so")
        with cr:
            av = st.selectbox("✅ 가용재고", cols, index=auto_idx(['가용재고']), key="sel_av")
            stk = st.selectbox("📦 정상재고", cols, index=auto_idx(['정상재고']), key="sel_stk")
            t3 = st.selectbox("🔥 3일 판매", cols, index=auto_idx(['3일'], exclude=['7일']), key="sel_t3")
            t7 = st.selectbox("📅 7일 판매", cols, index=auto_idx(['7일'], exclude=['3일']), key="sel_t7")
            reg = st.selectbox("📆 등록일", cols, index=auto_idx(['등록일']), key="sel_reg")

        s1, s2 = st.columns(2)
        lt = s1.number_input("⏳ 리드타임", value=7)
        ss = s2.number_input("🛡️ 안전재고", value=3)

        if st.button("📊 데이터 분석 시작", use_container_width=True, type="primary"):
            st.session_state.p = {'it':it, 'op':op, 'vn':vn, 'vi':vi, 'so':so, 'av':av, 'st':stk, 't3':t3, 't7':t7, 'reg':reg, 'lt':lt, 'ss':ss}
            st.session_state.analyzed = True
            st.rerun()

# ------------------------------------------------------------------
# [4단계: 데이터 편집 및 재고 관리] - 초과 입고 제로화 및 차감 기록
# ------------------------------------------------------------------
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📊 4단계: 데이터 편집 및 재고 관리")

    sh = get_sheet()
    p = st.session_state.p
    s_out, item, opt, vnd, v_it = p['so'], p['it'], p['op'], p['vn'], p['vi']
    stk, avl, t3, t7, lt, ss = p['st'], p['av'], p['t3'], p['t7'], p['lt'], p['ss']

    # UI 필터
    c1, c2, c3 = st.columns([1, 2, 1])
    with c1: f_mode = st.selectbox("🚦 상태 필터", ["전체보기", "정상만", "품절만"], index=1, key="v4_filter_mode")
    with c2: s_query = st.text_input("🔍 상품 검색", key="v4_search_query")
    with c3: s_date = st.date_input("🗓️ 입고 조회 날짜", datetime.now(KST).date(), key="v4_in_date")

    # 데이터 준비
    reorder_map, history_map = get_realtime_data_v4(s_date)
    df_work = st.session_state.df_raw.copy()

    for col in [stk, avl, t3, t7]:
        df_work[col] = pd.to_numeric(df_work[col], errors='coerce').fillna(0).astype(int)

    # 클린키 생성 (사장님 로직 유지)
    df_work['clean_key'] = df_work.apply(lambda r: super_clean(r[item]) + super_clean(r[opt]), axis=1)
    
    # 1. 리오더 잔량 계산 및 제로화 (핵심 기능)
    df_work["리오더 총합"] = df_work['clean_key'].map(reorder_map).fillna(0).astype(int)
    df_work["리오더 총합"] = df_work["리오더 총합"].clip(lower=0) 
    
    df_work["과거입고데이터"] = df_work['clean_key'].map(history_map).fillna(0).astype(int)
    df_work["리오더 입고"] = 0 
    
    # 2. 일판매 및 발주권장 계산
    df_work['일판매'] = df_work.apply(lambda r: int(round(r[t7]/7)) if r[t7]>0 else (int(round(r[t3]/3)) if r[t3]>0 else 0), axis=1)
    df_work['발주권장'] = ((df_work['일판매'] * (lt + ss)) - (df_work[avl] + df_work['리오더 총합'])).clip(lower=0).astype(int)

    # 필터링 적용
    is_so = df_work[s_out].astype(str).str.contains('품절', na=False)
    df_f = df_work[~is_so] if f_mode == "정상만" else (df_work[is_so] if f_mode == "품절만" else df_work)
    if s_query:
        df_f = df_f[df_f[item].astype(str).str.contains(s_query, case=False) | df_f[opt].astype(str).str.contains(s_query, case=False)]

    df_disp = df_f.rename(columns={s_out: "상태", vnd: "공급처", item: "상품명", opt: "옵션", v_it: "공급처상품명", stk: "정상재고", avl: "가용재고", t3: "3일발주"})
    
    with st.form("v4_storage_form", clear_on_submit=True):
        target_cols = ["상태", "공급처", "상품명", "옵션", "공급처상품명", "정상재고", "가용재고", "리오더 총합", "리오더 입고", "과거입고데이터", "3일발주", "일판매", "발주권장"]
        v4_ed = st.data_editor(
            df_disp[target_cols], use_container_width=True, hide_index=True, key="v4_main_editor",
            column_config={
                "리오더 총합": st.column_config.NumberColumn("📦 리오더잔량"),
                "리오더 입고": st.column_config.NumberColumn("📥 입고차감", min_value=0),
                "과거입고데이터": st.column_config.NumberColumn("📜 과거입고")
            }, 
            disabled=[c for c in target_cols if c != "리오더 입고"]
        )
        
        if st.form_submit_button("💾 입고 정보 저장 및 리오더 차감", use_container_width=True):
            edits = st.session_state.get("v4_main_editor", {}).get("edited_rows", {})
            if edits:
                try:
                    v7_sh = sh.worksheet("발주기록")
                    h_sh = sh.worksheet("입고기록")
                    t_date = s_date.strftime('%Y-%m-%d')
                    saved_count = 0
                    for r_idx, val in edits.items():
                        qty = int(val.get("리오더 입고", 0))
                        if qty > 0:
                            row = df_disp.iloc[int(r_idx)]
                            # 1. 발주기록 차감 저장
                            v7_sh.append_row([datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S'), str(row["상품명"]), str(row["옵션"]), str(row.get("공급처상품명", "")), 0, 0, -qty, 0, "입고차감", str(row.get("공급처", "미지정"))])
                            # 2. 입고기록 저장
                            h_sh.append_row([t_date, str(row["상품명"]), str(row["옵션"]), qty])
                            saved_count += 1
                    
                    if saved_count > 0:
                        st.success(f"✅ {saved_count}건 저장 완료!"); st.cache_data.clear(); time.sleep(1); st.rerun()
                except Exception as e: st.error(f"❌ 오류: {e}")

# ------------------------------------------------------------------
# [5단계: 최종 발주 요약] - 중복 합산 방지 + 메모 관리
# ------------------------------------------------------------------
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📋 5단계: 최종 발주 리스트 요약")

    reorder_map_v5, _ = get_realtime_data_v4(datetime.now(KST).date())
    df_v5 = st.session_state.df_raw.copy()
    df_v5['clean_key'] = df_v5.apply(lambda r: super_clean(r[item]) + super_clean(r[opt]), axis=1)

    for col in [stk, avl, t3, t7]:
        df_v5[col] = pd.to_numeric(df_v5[col], errors='coerce').fillna(0).astype(int)

    df_v5 = df_v5[~df_v5[s_out].astype(str).str.contains('품절', na=False)]
    if '메모' not in df_v5.columns: df_v5['메모'] = ""
    if 'add_order_dict' not in st.session_state: st.session_state.add_order_dict = {}

    df_v5["기존 리오더"] = df_v5['clean_key'].map(reorder_map_v5).fillna(0).astype(int)
    df_v5['추가발주입력'] = df_v5.index.map(st.session_state.add_order_dict).fillna(0).astype(int)
    df_v5['총 리오더 합계'] = df_v5["기존 리오더"] + df_v5['추가발주입력']

    df_v5['일판매'] = df_v5.apply(lambda r: int(round(r[t7]/7)) if r[t7]>0 else (int(round(r[t3]/3)) if r[t3]>0 else 0), axis=1)
    df_v5['발주권장'] = ((df_v5['일판매'] * (lt + ss)) - (df_v5[avl] + df_v5["기존 리오더"])).clip(lower=0).astype(int)

    f1, f2 = st.columns([1, 2])
    with f1: m5_f = st.selectbox("🚦 상태 필터", ["✅ 전체보기", "🚨 발주필요"], key="v5_f")
    with f2: s5_q = st.text_input("🔍 검색", key="v5_s")

    df_v5_v = df_v5.copy()
    if s5_q: df_v5_v = df_v5_v[df_v5_v[item].astype(str).str.contains(s5_q, case=False) | df_v5_v[opt].astype(str).str.contains(s5_q, case=False)]
    if m5_f == "🚨 발주필요": df_v5_v = df_v5_v[df_v5_v['발주권장'] > 0]
    
    with st.form("v5_comprehensive_form"):
        v_map = {item: "상품명", opt: "옵션", avl: "가용재고", "기존 리오더": "기존잔량", "추가발주입력": "추가발주", "총 리오더 합계": "총합계", "메모": "메모"}
        st.data_editor(
            df_v5_v[list(v_map.keys())].rename(columns=v_map), use_container_width=True, hide_index=True, key="v5_editor_final",
            column_config={"추가발주": st.column_config.NumberColumn("➕ 추가발주", min_value=0), "기존잔량": st.column_config.NumberColumn("📦 기존잔량", disabled=True), "총합계": st.column_config.NumberColumn("📊 총합계", disabled=True)}
        )
        if st.form_submit_button("✅ 1. 수량 확정 및 화면 갱신", use_container_width=True):
            edits = st.session_state.v5_editor_final.get("edited_rows", {})
            for r_idx, val in edits.items():
                idx = df_v5_v.index[int(r_idx)]
                if "추가발주" in val: st.session_state.add_order_dict[idx] = int(val["추가발주"])
                if "메모" in val: st.session_state.df_raw.at[idx, "메모"] = str(val["메모"])
            st.rerun()

    c_save, c_down = st.columns(2)
    with c_save:
        if st.button("💾 2. 구글 시트 최종 저장", use_container_width=True, type="primary"):
            raw_edits = st.session_state.v5_editor_final.get("edited_rows", {})
            if raw_edits:
                try:
                    ws_log = get_sheet().worksheet("발주기록")
                    now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
                    rows_to_add = []
                    for r_idx, changes in raw_edits.items():
                        if "추가발주" in changes and changes["추가발주"] > 0:
                            real_idx = df_v5_v.index[int(r_idx)]
                            rows_to_add.append([now_s, str(df_v5_v.at[real_idx, item]), str(df_v5_v.at[real_idx, opt]), str(df_v5_v.at[real_idx, v_it]), int(df_v5_v.at[real_idx, avl]), int(df_v5_v.at[real_idx, '기존 리오더']), int(changes["추가발주"]), int(df_v5_v.at[real_idx, '발주권장']), str(changes.get("메모", df_v5_v.at[real_idx, "메모"])), str(df_v5_v.at[real_idx, vnd])])
                    if rows_to_add:
                        ws_log.append_rows(rows_to_add); st.success("✅ 저장 완료!"); st.cache_data.clear(); st.session_state.add_order_dict = {}; time.sleep(1); st.rerun()
                except Exception as e: st.error(f"오류: {e}")

    with c_down:
        download_list = [k for k, v in st.session_state.add_order_dict.items() if v > 0]
        if download_list:
            df_down = df_v5[df_v5.index.isin(download_list)].copy()
            df_down['최종발주수량'] = df_down.index.map(st.session_state.add_order_dict)
            csv_data = df_down[[vnd, item, opt, v_it, '최종발주수량']].rename(columns={vnd: "공급처", item: "상품명", opt: "옵션", v_it: "공급처상품명", "최종발주수량": "수량"}).to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button("📥 3. 발주서(CSV) 다운로드", csv_data, f"발주서_{datetime.now(KST).strftime('%m%d_%H%M')}.csv", use_container_width=True)

# ------------------------------------------------------------------
# [6단계: 전체 히스토리 관리] - 회차 선택 기능 완벽 복구
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.subheader("📜 6단계: 전체 히스토리 관리")
    f1, f2, f3, f4 = st.columns([1.2, 0.8, 1.5, 1.5])
    with f1: d_range = st.date_input("🗓️ 1. 조회 범위", value=(datetime.now(KST).date(), datetime.now(KST).date()), key="v6_date_range")
    with f2: st.write(""); search_trigger = st.button("🔍 2. 내역 조회", use_container_width=True, type="primary")

    if search_trigger:
        try:
            worksheet = get_sheet().worksheet("발주기록")
            all_h = worksheet.get_all_values()
            if len(all_h) > 1:
                df_all = pd.DataFrame(all_h[1:], columns=["발주시간", "상품명", "옵션", "공급처상품명", "가용재고", "리오더잔량", "추가발주", "발주권장", "메모", "업체명"])
                df_all["날짜_만"] = df_all["발주시간"].str.slice(0, 10)
                s_d, e_d = d_range[0].strftime('%Y-%m-%d'), d_range[1].strftime('%Y-%m-%d')
                df_filtered = df_all[(df_all["날짜_만"] >= s_d) & (df_all["날짜_만"] <= e_d)].copy()
                st.session_state.v6_data = df_filtered
                st.session_state.v6_sessions = sorted(df_filtered["발주시간"].unique(), reverse=True)
                st.session_state.v6_display_text = f"🗓️ {s_d} ~ {e_d}"
        except Exception as e: st.error(f"오류: {e}")

    with f3: h_q = st.text_input("🔍 3. 상품명/옵션 검색", key="v6_search_q")
    with f4:
        if st.session_state.get('v6_sessions'):
            session_options = ["📊 선택 범위 전체 합산"] + [f"{len(st.session_state.v6_sessions)-i}회차 ({t[5:16]})" for i, t in enumerate(st.session_state.v6_sessions)]
            sel_session_label = st.selectbox("📦 4. 회차 선택", session_options, key="v6_session_select")
        else: sel_session_label = st.selectbox("📦 4. 회차 선택", ["조회 결과 없음"], disabled=True)

    if st.session_state.get('v6_data') is not None and sel_session_label:
        df_display = st.session_state.v6_data.copy()
        for col in ["가용재고", "리오더잔량", "추가발주", "발주권장"]: df_display[col] = pd.to_numeric(df_display[col], errors='coerce').fillna(0)
        
        if sel_session_label == "📊 선택 범위 전체 합산":
            df_display = df_display.groupby(["업체명", "상품명", "옵션", "공급처상품명"], as_index=False).agg({"발주시간": "max", "가용재고": "last", "리오더잔량": "last", "추가발주": "sum", "발주권장": "last", "메모": lambda x: " / ".join(set(filter(None, x.astype(str))))})
        else:
            target_time = st.session_state.v6_sessions[session_options.index(sel_session_label)-1]
            df_display = df_display[df_display["발주시간"] == target_time].copy()

        if h_q: df_display = df_display[df_display["상품명"].str.contains(h_q, case=False) | df_display["옵션"].str.contains(h_q, case=False)]
        st.dataframe(df_display[["발주시간", "업체명", "상품명", "옵션", "공급처상품명", "가용재고", "리오더잔량", "추가발주", "발주권장", "메모"]], use_container_width=True, hide_index=True)

# ------------------------------------------------------------------
# [7단계: 실시간 리오더 상황판] - 제로화 및 입고차감 메모 보정
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.subheader("🚀 7단계: 실시간 리오더 최종 잔량 상황판")
    try:
        ws_v7 = get_sheet().worksheet("발주기록")
        all_v7 = ws_v7.get_all_values()
        if len(all_v7) > 1:
            df_v7 = pd.DataFrame(all_v7[1:], columns=["발주시간", "상품명", "옵션", "공급처상품명", "가용재고", "기존리오더", "추가발주", "발주권장", "메모", "업체명"])
            df_v7["최종잔량"] = pd.to_numeric(df_v7["기존리오더"], errors='coerce').fillna(0) + pd.to_numeric(df_v7["추가발주"], errors='coerce').fillna(0)
            
            # 메모 보정 로직 (사장님 요청 반영)
            df_v7["메모"] = df_v7.apply(lambda r: f"{r['추가발주']}개 입고차감" if (to_int(r['추가발주']) < 0 and str(r['메모']).strip() == "입고차감") else r['메모'], axis=1)

            df_final = df_v7.groupby(["업체명", "상품명", "옵션", "공급처상품명"], as_index=False).agg({"발주시간": "max", "최종잔량": "sum", "메모": lambda x: " / ".join(dict.fromkeys(filter(None, x.astype(str))))})
            df_final["최종잔량"] = df_final["최종잔량"].clip(lower=0) # 제로화 핵심

            # 업체별 전광판
            df_v_sum = df_final.groupby("업체명")["최종잔량"].sum().reset_index().sort_values("최종잔량", ascending=False)
            df_v_sum = df_v_sum[df_v_sum["최종잔량"] > 0]
            v_cols = st.columns(4)
            for i, r in enumerate(df_v_sum.itertuples()):
                with v_cols[i % 4]: st.metric(r.업체명, f"{int(r.최종잔량):,} 개")
            
            st.write("#### 📋 상세 리스트")
            st.dataframe(df_final.sort_values("발주시간", ascending=False), use_container_width=True, hide_index=True)
    except Exception as e: st.error(f"7단계 오류: {e}")
