import streamlit as st
import pandas as pd
import numpy as np
import re
import time
import unicodedata
import gspread
import streamlit.components.v1 as components
from datetime import datetime, timedelta, timezone

# 1. 기본 설정 및 한국 시간
KST = timezone(timedelta(hours=9))
st.set_page_config(layout="wide", page_title="저스트원 재고관리 v4.0")

# --- [세션 상태 초기화] ---
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'analyzed' not in st.session_state: st.session_state.analyzed = False
if 'p' not in st.session_state: st.session_state.p = {}
if 'upload_key' not in st.session_state: st.session_state.upload_key = 0

# 새로고침 방지
components.html("<script>window.onbeforeunload = function() { return '변경사항이 저장되지 않을 수 있습니다.'; };</script>", height=0)

# --- [공통 보조 함수: 사장님 필수 로직] ---
def get_sheet():
    try:
        client = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        return client.open_by_key("1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg")
    except Exception as e:
        st.error(f"📡 시트 연결 실패: {e}")
        return None

def super_clean(t):
    if not t: return ""
    t = unicodedata.normalize('NFC', str(t))
    return re.sub(r'[^a-zA-Z0-9가-힣]', '', t).upper().strip()

def to_i(v):
    try: return int(float(str(v).replace(",", "").strip()))
    except: return 0

# [에러 해결] 실시간 데이터 호출 함수 복구
def get_realtime_data_v4(target_date):
    try:
        sh = get_sheet()
        if not sh: return {}, {}
        
        # 1. 발주기록 매핑 (리오더 + 추가리오더 합산)
        ws_v = sh.worksheet("발주기록")
        d_v = ws_v.get_all_values()
        r_map = {}
        if len(d_v) > 1:
            for row in d_v[1:]:
                try:
                    # 상품명(1) + 옵션(2) 키 생성
                    key = super_clean(row[1]) + super_clean(row[2])
                    val = to_i(row[5]) # 기존
                    add = to_i(row[6]) # 추가
                    if (val + add) != 0:
                        r_map[key] = r_map.get(key, 0) + (val + add)
                except: continue

        # 2. 입고기록 매핑 (해당 날짜 입고량)
        ws_h = sh.worksheet("입고기록")
        d_h = ws_h.get_all_values()
        h_map = {}
        t_str = target_date.strftime('%Y-%m-%d')
        if len(d_h) > 1:
            for row_h in d_h[1:]:
                try:
                    if t_str in str(row_h[0]):
                        h_key = super_clean(row_h[1]) + super_clean(row_h[2])
                        h_map[h_key] = h_map.get(h_key, 0) + to_i(row_h[3])
                except: continue
        return r_map, h_map
    except:
        return {}, {}

def sync_reorder_from_sheet(df_uploaded):
    # 이 함수는 get_realtime_data_v4를 사용하여 리오더 수량을 df에 합쳐줍니다.
    try:
        current_today = datetime.now(KST).date()
        r_map, _ = get_realtime_data_v4(current_today)
        
        if "리오더 수량" in df_uploaded.columns:
            df_uploaded = df_uploaded.drop(columns=["리오더 수량"])

        def final_match(r):
            u_key = super_clean(r.get('상품명', '')) + super_clean(r.get('옵션', ''))
            return r_map.get(u_key, 0)

        df_uploaded['리오더 수량'] = df_uploaded.apply(final_match, axis=1)
        st.success(f"✅ 리오더 합산 완료 (시트 데이터 {len(r_map)}건 로드)")
        return df_uploaded
    except Exception as e:
        st.error(f"동기화 오류: {e}")
        return df_uploaded

# --- [메인 UI 영역] ---
st.title("📦 저스트원 통합 재고 관리 v4.0")
tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 관리"])

with tab1:
    st.subheader("📁 1단계: 데이터 업로드")
    up_file = st.file_uploader("파일 업로드", type=['xlsx', 'xls', 'csv'], key=f"up_{st.session_state.upload_key}")

    if st.button("🔄 화면 전체 초기화", use_container_width=True):
        for key in list(st.session_state.keys()):
            if key != "upload_key": del st.session_state[key]
        st.session_state.upload_key += 1
        st.rerun()

    if up_file and st.session_state.df_raw is None:
        try:
            df = pd.read_csv(up_file) if up_file.name.endswith('.csv') else pd.read_excel(up_file)
            df.columns = [str(c).strip() for c in df.columns]
            with st.spinner("🔄 구글 시트 데이터(기존+추가) 불러오는 중..."):
                df = sync_reorder_from_sheet(df)
            st.session_state.df_raw = df.fillna("")
            st.rerun()
        except Exception as e:
            st.error(f"파일 로드 오류: {e}")

    if st.session_state.df_raw is not None:
        st.divider()
        st.subheader("📋 2단계: 매핑 항목 확인")
        cols = st.session_state.df_raw.columns.tolist()
        
        def auto_idx(keys, exclude_keys=None):
            for i, c in enumerate(cols):
                col_n = str(c)
                if exclude_keys and any(ek in col_n for ek in exclude_keys): continue
                if any(k in col_n for k in keys): return i
            return 0

        c1, c2 = st.columns(2)
        with c1:
            it = st.selectbox("📦 상품명", cols, index=auto_idx(['상품명']), key="sel_it")
            op = st.selectbox("🎨 옵션", cols, index=auto_idx(['옵션']), key="sel_op")
            vn = st.selectbox("🏭 공급처", cols, index=auto_idx(['공급처']), key="sel_vn")
            vi = st.selectbox("🆔 공급처 상품명", cols, index=auto_idx(['공급처상품명', '공급처상품']), key="sel_vi")
            so = st.selectbox("🚫 품절여부", cols, index=auto_idx(['품절']), key="sel_so")
        with c2:
            av = st.selectbox("✅ 가용재고", cols, index=auto_idx(['가용재고']), key="sel_av")
            stk = st.selectbox("📦 정상재고", cols, index=auto_idx(['정상재고']), key="sel_stk")
            t3 = st.selectbox("🔥 3일 판매", cols, index=auto_idx(['3일'], exclude_keys=['7일', '품절']), key="sel_t3")
            t7 = st.selectbox("📅 7일 판매", cols, index=auto_idx(['7일', '1주'], exclude_keys=['3일']), key="sel_t7")
            reg = st.selectbox("📆 등록일", cols, index=auto_idx(['등록일']), key="sel_reg")

        st.divider()
        st.subheader("🚀 3단계: 분석 설정")
        s1, s2 = st.columns(2)
        with s1: lt_val = st.number_input("⏳ 리드타임(LT)", value=7, key="inp_lt")
        with s2: ss_val = st.number_input("🛡️ 안전재고(SS)", value=3, key="inp_ss")

        if st.button("📊 데이터 분석 시작", type="primary", use_container_width=True):
            st.session_state.p = {
                'it': it, 'op': op, 'vn': vn, 'vi': vi, 'so': so, 
                'av': av, 'st': stk, 't3': t3, 't7': t7, 'reg': reg,
                'lt': lt_val, 'ss': ss_val
            }
            df_final = st.session_state.df_raw.copy()
            for col in [t3, t7, '리오더 수량']:
                df_final[col] = pd.to_numeric(df_final.get(col, 0), errors='coerce').fillna(0)
            
            st.session_state.df_raw = df_final
            st.session_state.analyzed = True
            st.rerun()


# ------------------------------------------------------------------
# [4단계: 데이터 편집 및 재고 관리] - 초과 입고 제로화 및 10개 컬럼 동기화
# ------------------------------------------------------------------
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📊 4단계: 데이터 편집 및 재고 관리")

    sh = get_sheet()
    p = st.session_state.p
    s_out, item, opt = p['so'], p['it'], p['op']
    vnd, v_it = p['vn'], p['vi']
    stk, avl, t3, t7 = p['st'], p['av'], p['t3'], p['t7']
    lt, ss = p['lt'], p['ss']

    # UI 필터
    c1, c2, c3 = st.columns([1, 2, 1])
    with c1: f_mode = st.selectbox("🚦 상태 필터", ["전체보기", "정상만", "품절만"], index=1, key="v4_filter_mode")
    with c2: s_query = st.text_input("🔍 상품 검색", key="v4_search_query")
    with c3: s_date = st.date_input("🗓️ 입고 조회 날짜", datetime.now(KST).date(), key="v4_in_date")

    # [1] 데이터 준비 및 실시간 리오더 로드
    reorder_map, history_map = get_realtime_data_v4(s_date)
    df_work = st.session_state.df_raw.copy()

    for col in [stk, avl, t3, t7]:
        df_work[col] = pd.to_numeric(df_work[col], errors='coerce').fillna(0).astype(int)

    # 상품 식별 키 생성 함수 (NFC 정규화 포함)
    def get_clean_key_v4(r):
        import unicodedata, re
        n = re.sub(r'[^a-zA-Z0-9가-힣]', '', unicodedata.normalize('NFC', str(r[item]))).upper().strip()
        o = re.sub(r'[^a-zA-Z0-9가-힣]', '', unicodedata.normalize('NFC', str(r[opt]))).upper().strip()
        return n + o

    df_work['clean_key'] = df_work.apply(get_clean_key_v4, axis=1)
    
    # 리오더 잔량 계산 및 ⭐제로화 적용
    df_work["리오더 총합"] = df_work['clean_key'].map(reorder_map).fillna(0).astype(int)
    df_work["리오더 총합"] = df_work["리오더 총합"].clip(lower=0) # 마이너스 방지
    
    df_work["과거입고데이터"] = df_work['clean_key'].map(history_map).fillna(0).astype(int)
    df_work["리오더 입고"] = 0 
    
    # 일판매 및 발주권장 재계산
    df_work['일판매'] = df_work.apply(lambda r: int(round(r[t7]/7)) if r[t7]>0 else (int(round(r[t3]/3)) if r[t3]>0 else 0), axis=1)
    df_work['발주권장'] = ((df_work['일판매'] * (lt + ss)) - (df_work[avl] + df_work['리오더 총합'])).clip(lower=0).astype(int)

    # 필터 적용
    is_so = df_work[s_out].astype(str).str.contains('품절', na=False)
    df_f = df_work[~is_so] if f_mode == "정상만" else (df_work[is_so] if f_mode == "품절만" else df_work)
    if s_query:
        df_f = df_f[df_f[item].astype(str).str.contains(s_query, case=False) | df_f[opt].astype(str).str.contains(s_query, case=False)]

    # 화면 표시용 컬럼명 변경
    df_disp = df_f.rename(columns={s_out: "상태", vnd: "공급처", item: "상품명", opt: "옵션", v_it: "공급처상품명", stk: "정상재고", avl: "가용재고", t3: "3일발주"})
    
    # [2] 에디터 및 입고 차감 폼
    with st.form("v4_storage_form", clear_on_submit=True):
        target_cols = ["상태", "공급처", "상품명", "옵션", "공급처상품명", "정상재고", "가용재고", "리오더 총합", "리오더 입고", "과거입고데이터", "발주권장"]
        
        v4_ed = st.data_editor(
            df_disp[target_cols], 
            use_container_width=True, 
            hide_index=True, 
            key="v4_main_editor", 
            column_config={
                "리오더 총합": st.column_config.NumberColumn("📦 현재잔량", disabled=True),
                "리오더 입고": st.column_config.NumberColumn("📥 입고수량(차감)", min_value=0, help="입고된 수량을 입력하면 잔량에서 차감됩니다."),
                "과거입고데이터": st.column_config.NumberColumn("📜 입고기록", disabled=True),
                "발주권장": st.column_config.NumberColumn("💡 권장", disabled=True)
            }, 
            disabled=[c for c in target_cols if c != "리오더 입고"]
        )
        
        submit_btn = st.form_submit_button("💾 입고 정보 반영 및 리오더 차감", use_container_width=True)
        
        if submit_btn:
            # 에디터에서 수정된 행(입고 수량이 입력된 행) 추출
            edits = st.session_state.get("v4_main_editor", {}).get("edited_rows", {})
            
            if edits:
                try:
                    v7_sh = sh.worksheet("발주기록")
                    h_sh = sh.worksheet("입고기록")
                    now_ts = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
                    t_date = s_date.strftime('%Y-%m-%d')
                    
                    rows_to_v7 = [] # 발주기록 시트용 (차감 내역)
                    rows_to_h = []  # 입고기록 시트용 (단순 기록)
                    
                    for r_idx, val in edits.items():
                        qty = int(val.get("리오더 입고", 0))
                        if qty > 0:
                            row = df_disp.iloc[int(r_idx)]
                            
                            # ⭐ [핵심] 발주기록 10개 컬럼 구조 그대로 저장 (6,7단계와 호환)
                            # 컬럼: 발주시간, 상품명, 옵션, 공급처상품명, 가용재고, 리오더잔량(0), 추가발주(-qty), 발주권장, 메모, 업체명
                            rows_to_v7.append([
                                now_ts, 
                                str(row["상품명"]), 
                                str(row["옵션"]), 
                                str(row.get("공급처상품명", "")), 
                                int(row["가용재고"]), 
                                0,              # 기존 리오더 (차감 시에는 0 처리)
                                -qty,           # 추가발주 열에 마이너스 입력하여 차감
                                int(row["발주권장"]), 
                                "입고차감",      # 메모에 입고차감 명시
                                str(row.get("공급처", "미지정"))
                            ])
                            
                            # 입고기록 전용 시트 저장용
                            rows_to_h.append([t_date, str(row["상품명"]), str(row["옵션"]), qty])
                    
                    if rows_to_v7:
                        v7_sh.append_rows(rows_to_v7)
                        h_sh.append_rows(rows_to_h)
                        st.success(f"✅ {len(rows_to_v7)}건의 입고 차감이 완료되었습니다!")
                        st.cache_data.clear() 
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.warning("입력된 수량이 없습니다.")
                except Exception as e:
                    st.error(f"❌ 저장 중 시트 오류 발생: {e}")
            else:
                st.warning("⚠️ 입고 수량을 입력한 뒤 버튼을 눌러주세요.")


                

# ------------------------------------------------------------------
# [5단계: 최종 발주 요약] - 저장 로직 중복 방지 및 10개 컬럼 동기화
# ------------------------------------------------------------------
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📋 5단계: 최종 발주 리스트 요약")

    # [1] 실시간 잔량 로드 및 데이터 준비
    # 실시간으로 시트에서 현재 리오더 잔량을 가져와 뻥튀기 방지
    reorder_map_v5, _ = get_realtime_data_v4(datetime.now(KST).date())
    df_v5 = st.session_state.df_raw.copy()

    # 상품 식별 키 생성 함수 (NFC 정규화)
    def get_clean_key_v5(r):
        import unicodedata, re
        n = re.sub(r'[^a-zA-Z0-9가-힣]', '', unicodedata.normalize('NFC', str(r[item]))).upper().strip()
        o = re.sub(r'[^a-zA-Z0-9가-힣]', '', unicodedata.normalize('NFC', str(r[opt]))).upper().strip()
        return n + o

    df_v5['clean_key'] = df_v5.apply(get_clean_key_v5, axis=1)
    for col in [stk, avl, t3, t7]:
        df_v5[col] = pd.to_numeric(df_v5[col], errors='coerce').fillna(0).astype(int)

    # 품절 제외 및 메모/추가발주 딕셔너리 초기화
    df_v5 = df_v5[~df_v5[s_out].astype(str).str.contains('품절', na=False)]
    if '메모' not in df_v5.columns: df_v5['메모'] = ""
    if 'add_order_dict' not in st.session_state: st.session_state.add_order_dict = {}

    # 수치 계산 (기존 리오더는 표시용으로만 사용)
    df_v5["기존 리오더"] = df_v5['clean_key'].map(reorder_map_v5).fillna(0).astype(int)
    # 초과입고 제로화 적용 (표시 상 마이너스 방지)
    df_v5["기존 리오더"] = df_v5["기존 리오더"].clip(lower=0)
    
    # 세션에 임시 저장된 추가발주 수량 매핑
    df_v5['추가발주입력'] = df_v5.index.map(st.session_state.add_order_dict).fillna(0).astype(int)
    df_v5['총 리오더 합계'] = df_v5["기존 리오더"] + df_v5['추가발주입력']

    # 발주 권장 계산 (현재 잔량 기준)
    df_v5['일판매'] = df_v5.apply(lambda r: int(round(r[t7]/7)) if r[t7]>0 else (int(round(r[t3]/3)) if r[t3]>0 else 0), axis=1)
    df_v5['발주권장'] = ((df_v5['일판매'] * (lt + ss)) - (df_v5[avl] + df_v5["기존 리오더"])).clip(lower=0).astype(int)

    # [2] 필터 영역
    f1, f2 = st.columns([1, 2])
    with f1: m5_f = st.selectbox("🚦 상태 필터", ["✅ 전체보기", "🚨 발주필요"], key="v5_f")
    with f2: s5_q = st.text_input("🔍 검색", key="v5_s")

    df_v5_v = df_v5.copy()
    if s5_q:
        df_v5_v = df_v5_v[df_v5_v[item].astype(str).str.contains(s5_q, case=False) | df_v5_v[opt].astype(str).str.contains(s5_q, case=False)]
    if m5_f == "🚨 발주필요": 
        df_v5_v = df_v5_v[df_v5_v['발주권장'] > 0]
    
    v_map = {
        item: "상품명", opt: "옵션", avl: "가용재고", 
        "기존 리오더": "기존잔량", "추가발주입력": "추가발주", "총 리오더 합계": "총합계", "메모": "메모"
    }
    actual_cols = [c for c in v_map.keys() if c in df_v5_v.columns]

    # [3] 에디터 및 수량 확정 (폼 사용)
    with st.form("v5_form"):
        df_ed = df_v5_v[actual_cols].rename(columns=v_map)
        edited_df = st.data_editor(
            df_ed, use_container_width=True, hide_index=True, key="v5_editor_final",
            column_config={
                "추가발주": st.column_config.NumberColumn("➕ 추가발주", min_value=0, help="이번에 새롭게 발주할 수량을 입력하세요."),
                "기존잔량": st.column_config.NumberColumn("📦 기존잔량", disabled=True),
                "총합계": st.column_config.NumberColumn("📊 최종예상", disabled=True),
            }
        )
        
        confirm_btn = st.form_submit_button("✅ 1. 입력 수량 확정 (저장 전 필수)", use_container_width=True)
        
        if confirm_btn:
            for idx, row in edited_df.iterrows():
                real_idx = df_v5_v.index[idx]
                st.session_state.add_order_dict[real_idx] = int(row["추가발주"])
                st.session_state.df_raw.at[real_idx, "메모"] = str(row["메모"])
            st.success("수량이 확정되었습니다. '2. 구글 시트 최종 저장' 버튼을 누르면 시트에 반영됩니다.")
            st.rerun()

    # [4] 저장 및 다운로드
    st.write("")
    c_save, c_down = st.columns(2)
    
    with c_save:
        if st.button("💾 2. 구글 시트 최종 저장", use_container_width=True, type="primary"):
            rows_to_add = []
            now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
            
            # 확정된 수량 중 0보다 큰 것들만 골라 시트에 기록
            for idx, add_qty in st.session_state.add_order_dict.items():
                if add_qty > 0:
                    # ⭐ 따블 방지 핵심: 기존 리오더 열에는 0을 넣고, 추가발주 열에만 수량을 넣음
                    # 이렇게 해야 7단계 상황판에서 (이전 합계) + (새로운 추가분) = (정확한 총합)이 됨
                    rows_to_add.append([
                        now_s, 
                        str(df_v5.at[idx, item]), 
                        str(df_v5.at[idx, opt]), 
                        str(df_v5.at[idx, v_it]), 
                        int(df_v5.at[idx, avl]), 
                        0,               # F열(기존리오더): 중복 합산 방지를 위해 0으로 저장
                        int(add_qty),    # G열(추가발주): 이번 순수 추가분만 저장
                        int(df_v5.at[idx, '발주권장']), 
                        str(df_v5.at[idx, "메모"]),
                        str(df_v5.at[idx, vnd])
                    ])
            
            if rows_to_add:
                try:
                    ws_log = get_sheet().worksheet("발주기록")
                    ws_log.append_rows(rows_to_add)
                    st.success(f"✅ {len(rows_to_add)}건 저장 완료! 수량이 정확하게 합산되었습니다.")
                    # 저장 후 세션/캐시 비우기
                    st.session_state.add_order_dict = {}
                    st.cache_data.clear()
                    time.sleep(1)
                    st.rerun()
                except Exception as e:
                    if "429" in str(e):
                        st.error("🚨 구글 API 호출 한도 초과! 1분만 기다렸다가 다시 시도해주세요.")
                    else:
                        st.error(f"❌ 시트 저장 실패: {e}")
            else:
                st.warning("⚠️ 저장할 수량이 없습니다. 1번 확정 버튼을 먼저 눌러주세요.")

    with c_down:
        # 다운로드용 데이터 준비
        df_down = df_v5.copy()
        df_down['최종발주수량'] = df_down.index.map(st.session_state.add_order_dict).fillna(0).astype(int)
        df_final_down = df_down[df_down['최종발주수량'] > 0].copy()
        
        if not df_final_down.empty:
            csv_data = df_final_down[[vnd, item, opt, v_it, '최종발주수량']].rename(columns={
                vnd: "공급처", item: "상품명", opt: "옵션", v_it: "공급처상품명", "최종발주수량": "수량"
            })
            st.download_button(
                label="📥 3. 발주서(CSV) 다운로드",
                data=csv_data.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig'),
                file_name=f"발주서_{datetime.now(KST).strftime('%m%d_%H%M')}.csv",
                mime="text/csv",
                use_container_width=True
            )
        else:
            st.button("📥 3. 다운로드 (입력값 없음)", disabled=True, use_container_width=True)



# ------------------------------------------------------------------
# [6단계: 전체 히스토리 관리] - 5단계 저장 구조(10개 컬럼) 완벽 반영
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.subheader("📜 6단계: 전체 히스토리 관리")

    f1, f2, f3, f4 = st.columns([1.2, 0.8, 1.5, 1.5])
    
    with f1:
        today = datetime.now(KST).date()
        # 🗓️ 조회 범위 (에러 방지를 위해 오늘 날짜 기본값 세팅)
        d_range = st.date_input("🗓️ 1. 조회 범위", value=(today, today), key="v6_date_range")
    
    with f2:
        st.write(""); st.write("") 
        search_trigger = st.button("🔍 2. 내역 조회", use_container_width=True, type="primary")

    # 세션 상태 초기화
    if 'v6_data' not in st.session_state: st.session_state.v6_data = None
    if 'v6_sessions' not in st.session_state: st.session_state.v6_sessions = []
    if 'v6_display_text' not in st.session_state: st.session_state.v6_display_text = ""

    # [내역 조회 로직]
    if search_trigger:
        try:
            with st.spinner("📡 발주 내역을 불러오는 중..."):
                worksheet = get_sheet().worksheet("발주기록")
                all_h = worksheet.get_all_values()
                if len(all_h) > 1:
                    df_all = pd.DataFrame(all_h[1:])
                    
                    # 5단계 저장 로직의 10개 컬럼 순서와 완벽 일치 (순서 절대 유지)
                    cols_list = ["발주시간", "상품명", "옵션", "공급처상품명", "가용재고", "리오더잔량", "추가발주", "발주권장", "메모", "업체명"]
                    df_all.columns = cols_list[:df_all.shape[1]]
                    
                    # 날짜/시간 전처리
                    df_all["날짜_만"] = df_all["발주시간"].astype(str).str.slice(0, 10)
                    
                    # 날짜 필터링 (시작일/종료일 안전하게 추출)
                    if isinstance(d_range, (list, tuple)) and len(d_range) == 2:
                        s_d, e_d = d_range[0].strftime('%Y-%m-%d'), d_range[1].strftime('%Y-%m-%d')
                        st.session_state.v6_display_text = f"🗓️ {s_d} ~ {e_d}"
                    else:
                        # 날짜가 하나만 선택되었을 때의 처리
                        s_d = d_range[0].strftime('%Y-%m-%d') if isinstance(d_range, (list, tuple)) else d_range.strftime('%Y-%m-%d')
                        e_d = s_d
                        st.session_state.v6_display_text = f"🗓️ {s_d} 당일"
                    
                    df_filtered = df_all[(df_all["날짜_만"] >= s_d) & (df_all["날짜_만"] <= e_d)].copy()
                    
                    if not df_filtered.empty:
                        st.session_state.v6_data = df_filtered
                        st.session_state.v6_sessions = sorted(df_filtered["발주시간"].unique(), reverse=True)
                        st.success(f"✅ {len(df_filtered)}건의 내역을 찾았습니다.")
                    else:
                        st.session_state.v6_data = None
                        st.warning("💡 해당 기간에 저장된 내역이 없습니다.")
                else:
                    st.info("💡 시트가 비어있습니다.")
        except Exception as e:
            st.error(f"📡 데이터 로드 오류: {e}")

    # [필터 및 회차 선택 UI]
    with f3: h_q = st.text_input("🔍 3. 상품명/옵션 검색", key="v6_search_q")
    with f4:
        if st.session_state.v6_sessions:
            # 회차 선택 옵션 구성
            session_options = ["📊 선택 범위 전체 합산"] + [f"{len(st.session_state.v6_sessions)-i}회차 ({t[5:16]})" for i, t in enumerate(st.session_state.v6_sessions)]
            sel_session_label = st.selectbox("📦 4. 회차 선택", session_options, key="v6_session_select")
        else:
            st.selectbox("📦 4. 회차 선택", ["조회 결과 없음"], disabled=True)
            sel_session_label = None

    # [데이터 출력창]
    if st.session_state.v6_data is not None and sel_session_label:
        df_display = st.session_state.v6_data.copy()
        
        # 숫자 컬럼 강제 변환
        num_cols = ["가용재고", "리오더잔량", "추가발주", "발주권장"]
        for col in num_cols:
            if col in df_display.columns:
                df_display[col] = pd.to_numeric(df_display[col], errors='coerce').fillna(0).astype(int)

        # 회차 필터링 및 합산 로직
        if sel_session_label == "📊 선택 범위 전체 합산":
            display_title = st.session_state.v6_display_text + " 발주 합계"
            # 합산 시 업체명 포함 필수
            df_display = df_display.groupby(["업체명", "상품명", "옵션", "공급처상품명"], as_index=False).agg({
                "발주시간": "max", 
                "가용재고": "last",
                "리오더잔량": "last",
                "추가발주": "sum",
                "발주권장": "last",
                "메모": lambda x: " / ".join(dict.fromkeys(filter(None, x.astype(str))))
            })
        else:
            # 특정 회차 선택 시
            try:
                # 라벨 인덱스를 통해 실제 발주시간 찾기
                target_idx = session_options.index(sel_session_label) - 1
                target_time = st.session_state.v6_sessions[target_idx]
                df_display = df_display[df_display["발주시간"] == target_time].copy()
                display_title = f"✅ {sel_session_label} 상세 내역"
            except:
                display_title = "조회 오류"

        # 검색어 필터 적용
        if h_q:
            df_display = df_display[
                df_display["상품명"].astype(str).str.contains(h_q, case=False) | 
                df_display["옵션"].astype(str).str.contains(h_q, case=False)
            ]

        # 최종 화면 출력
        if not df_display.empty:
            st.write(f"#### {display_title}")
            # 사장님이 원하시는 10개 컬럼 순서 재정렬
            view_order = ["발주시간", "업체명", "상품명", "옵션", "공급처상품명", "가용재고", "리오더잔량", "추가발주", "발주권장", "메모"]
            # 실제 존재하는 컬럼만 필터링 (에러 방지)
            final_cols = [c for c in view_order if c in df_display.columns]
            
            st.dataframe(df_display[final_cols], use_container_width=True, hide_index=True)
            
            # CSV 다운로드 버튼
            csv_data = df_display[final_cols].to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button(
                label=f"📥 {display_title} CSV 다운로드", 
                data=csv_data, 
                file_name=f"발주히스토리_{datetime.now(KST).strftime('%m%d_%H%M')}.csv", 
                mime="text/csv",
                use_container_width=True
            )
        else:
            st.info("💡 검색 조건에 맞는 내역이 없습니다.")

        

# ------------------------------------------------------------------
# [7단계: 실시간 리오더 최종 잔량 상황판] - 제로화 및 메모 보정 통합
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.subheader("🚀 7단계: 실시간 리오더 최종 잔량 상황판")

    # [1] 데이터 캐싱 함수 (실시간 업데이트 대응)
    @st.cache_data(ttl=3600)
    def get_v7_data_cached():
        try:
            ws_v7 = get_sheet().worksheet("발주기록")
            return ws_v7.get_all_values()
        except Exception as e:
            st.error(f"시트 로드 실패: {e}")
            return []

    try:
        all_v7 = get_v7_data_cached()
        if len(all_v7) > 1:
            # 컬럼 정의 (시트 구조에 맞게 매핑)
            df_v7 = pd.DataFrame(all_v7[1:])
            # 데이터 컬럼 개수가 부족할 경우를 대비해 슬라이싱 처리
            cols_list = ["발주시간","업체명", "상품명", "옵션", "공급처상품명", "가용재고", "기존리오더", "추가발주", "발주권장", "메모"]
            df_v7.columns = cols_list[:df_v7.shape[1]]
            
            # 수치 변환 및 기본 전처리
            df_v7["기존리오더"] = pd.to_numeric(df_v7["기존리오더"], errors='coerce').fillna(0).astype(int)
            df_v7["추가발주"] = pd.to_numeric(df_v7["추가발주"], errors='coerce').fillna(0).astype(int)
            df_v7["최종잔량"] = df_v7["기존리오더"] + df_v7["추가발주"]
            df_v7["날짜_순수"] = df_v7["발주시간"].str.slice(0, 10)

            # [2] 메모 보정: "입고차감"만 있을 때 수량 정보를 붙여 가독성 증대
            def fix_memo_v7(row):
                m = str(row.get('메모', '')).strip()
                q = row.get('추가발주', 0)
                if q < 0 and (m == "입고차감" or m == ""):
                    return f"{abs(q)}개 입고차감"
                return m
            df_v7["메모"] = df_v7.apply(fix_memo_v7, axis=1)

            # [3] 필터 UI 영역 (날짜, 검색, 업체 선택)
            f1, f2, f3, f4 = st.columns([1.2, 0.6, 1.5, 1.2])
            
            with f1:
                # 날짜 범위 검색 (기억 안 나신다던 그 부분!)
                start_d = (datetime.now(KST) - timedelta(days=30)).date()
                end_d = datetime.now(KST).date()
                d_range = st.date_input("🗓️ 기간 조회", value=(start_d, end_d), key="v7_dr")
            
            with f2:
                st.write(""); st.write("")
                if st.button("🔄 업데이트", use_container_width=True, type="primary"):
                    st.cache_data.clear()
                    st.rerun()
            
            with f3:
                q_v7 = st.text_input("🔍 상품명/옵션 검색", key="v7_qs")
            
            with f4:
                v_list = ["전체 업체"] + sorted(df_v7["업체명"].unique().tolist()) if "업체명" in df_v7.columns else ["전체 업체"]
                v_choice = st.selectbox("🏭 업체 필터", v_list, key="v7_vs")

            # [4] 필터링 로직 적용
            df_f = df_v7.copy()
            
            # 날짜 필터 (범위 선택 대응)
            if isinstance(d_range, (list, tuple)) and len(d_range) == 2:
                s_str, e_str = d_range[0].strftime('%Y-%m-%d'), d_range[1].strftime('%Y-%m-%d')
                df_f = df_f[(df_f["날짜_순수"] >= s_str) & (df_f["날짜_순수"] <= e_str)]
            
            # 업체 필터
            if v_choice != "전체 업체":
                df_f = df_f[df_f["업체명"] == v_choice]
            
            # 검색어 필터
            if q_v7:
                df_f = df_f[df_f["상품명"].str.contains(q_v7, case=False) | df_f["옵션"].str.contains(q_v7, case=False)]

            if not df_f.empty:
                # [5] 데이터 그룹화 및 최종 합산
                group_cols = ["업체명", "상품명", "옵션", "공급처상품명"]
                df_final = df_f.groupby(group_cols, as_index=False).agg({
                    "발주시간": "max", 
                    "최종잔량": "sum",
                    "메모": lambda x: " / ".join(dict.fromkeys(filter(None, x.astype(str))))
                })

                # ⭐ [핵심] 최종 합산 잔량 제로화 (초과 입고 시 마이너스 방지)
                df_final["최종잔량"] = df_final["최종잔량"].clip(lower=0)
                
                # 잔량이 남은 것 위주로 정렬
                df_final = df_final.sort_values(["최종잔량", "발주시간"], ascending=[False, False])

                # [6] 상단 전광판 (메트릭)
                st.write("### 📊 업체별 미입고 현황")
                df_v_sum = df_final.groupby("업체명")["최종잔량"].sum().reset_index()
                df_v_sum = df_v_sum[df_v_sum["최종잔량"] > 0].sort_values("최종잔량", ascending=False)
                
                if not df_v_sum.empty:
                    v_cols = st.columns(4)
                    for i, r in enumerate(df_v_sum.itertuples()):
                        with v_cols[i % 4]:
                            st.metric(label=r.업체명, value=f"{int(r.최종잔량):,} 개")
                
                # [7] 상세 리스트 출력
                st.write("#### 📋 상세 리스트 (미입고 품목)")
                st.dataframe(
                    df_final[df_final["최종잔량"] > 0], # 잔량 있는 것만 노출
                    use_container_width=True, 
                    hide_index=True,
                    column_config={
                        "발주시간": st.column_config.TextColumn("🕒 최종작업일"),
                        "최종잔량": st.column_config.NumberColumn("🔢 미입고잔량", format="%d"), 
                        "메모": st.column_config.TextColumn("📝 히스토리/비고", width="large")
                    }
                )
            else:
                st.info("💡 조회된 조건에 맞는 데이터가 없습니다.")
        else:
            st.info("📅 아직 기록된 발주 내역이 없습니다.")
            
    except Exception as e:
        st.error(f"❌ 7단계 화면 구성 중 오류 발생: {e}")
