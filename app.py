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
# [5단계: 최종 발주 요약] - 첫 로딩 시 "긴급발주" 필터 기본값 설정
# ------------------------------------------------------------------
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📋 5단계: 최종 발주 요약 (🚨 긴급 항목 우선 로딩)")

    # [1] 데이터 준비 및 숫자 변환
    df_v5 = st.session_state.df_raw.copy()
    reorder_map_v5, _ = get_realtime_data_v4(datetime.now(KST).date())

    for col in [stk, avl, t3, t7]:
        if col in df_v5.columns:
            df_v5[col] = pd.to_numeric(df_v5[col], errors='coerce').fillna(0).astype(int)

    # [2] 긴급 상품 판별 및 상태 부여 (옵션 묶음 로직)
    def get_clean_key_v5(r):
        import unicodedata, re
        n = re.sub(r'[^a-zA-Z0-9가-힣]', '', unicodedata.normalize('NFC', str(r.get(item, "")))).upper().strip()
        o = re.sub(r'[^a-zA-Z0-9가-힣]', '', unicodedata.normalize('NFC', str(r.get(opt, "")))).upper().strip()
        return n + o
    df_v5['clean_key'] = df_v5.apply(get_clean_key_v5, axis=1)

    df_v5["리오더잔량"] = df_v5['clean_key'].map(reorder_map_v5).fillna(0).astype(int)
    df_v5['일판매'] = df_v5.apply(lambda r: int(round(r[t7]/7)) if r[t7]>0 else (int(round(r[t3]/3)) if r[t3]>0 else 0), axis=1)
    df_v5['발주권장'] = ((df_v5['일판매'] * (lt + ss)) - (df_v5[avl] + df_v5["리오더잔량"])).clip(lower=0).astype(int)

    # ⭐ 상품 묶음 알림: 하나라도 부족하면 해당 상품명(item) 전체를 긴급으로!
    urgent_item_names = df_v5[df_v5['발주권장'] > 0][item].unique()
    
    def set_status_logic(row):
        if row[item] in urgent_item_names:
            return "🚨 긴급발주", 0
        return "✅ 정상", 1
    
    df_v5[['알림표기', 'sort_order']] = df_v5.apply(lambda x: pd.Series(set_status_logic(x)), axis=1)

    # [3] 필터 UI - ⭐ index=1 설정을 통해 "🚨 긴급발주!
    f1, f2 = st.columns([1, 2])
    with f1: 
        m5_f = st.selectbox(
            "🚦 상태 필터", 
            ["전체보기", "🚨 긴급발주 상품", "✅ 정상 상품"], 
            index=1,  # 👈 앱 열자마자 2번째 항목(긴급)이 선택되도록 고정
            key="v5_auto_filter"
        )
    with f2: 
        s5_q = st.text_input("🔍 검색 (상품명/옵션)", key="v5_auto_search")

    # 필터 적용 로직
    df_v5_v = df_v5.copy()
    if s5_q:
        df_v5_v = df_v5_v[df_v5_v[item].astype(str).str.contains(s5_q, case=False) | df_v5_v[opt].astype(str).str.contains(s5_q, case=False)]
    
    if m5_f == "🚨 긴급발주 상품":
        df_v5_v = df_v5_v[df_v5_v['알림표기'] == "🚨 긴급발주"]
    elif m5_f == "✅ 정상 상품":
        df_v5_v = df_v5_v[df_v5_v['알림표기'] == "✅ 정상"]

    # 정렬: 긴급이 위로, 그 안에서 상품/옵션순
    df_v5_v = df_v5_v.sort_values(by=['sort_order', item, opt])

    # [4] 에디터 및 순서 고정 (KeyError 방지 포함)
    with st.form("v5_loading_fix_form"):
        target_order = ["알림표기", vnd, item, opt, v_it, avl, "리오더잔량", "추가발주입력", "총합계", "발주권장", "메모"]
        v_map = {"알림표기": "상태", vnd: "공급처", item: "상품명", opt: "옵션", v_it: "공급처상품명",
                 avl: "가용재고", "리오더잔량": "기존잔량", "추가발주입력": "추가발주", "총합계": "총합계", "발주권장": "권장발주", "메모": "메모"}

        # 화면용 데이터 최종 정비
        df_v5_v['추가발주입력'] = df_v5_v.index.map(st.session_state.add_order_dict).fillna(0).astype(int)
        df_v5_v['총합계'] = df_v5_v["리오더잔량"] + df_v5_v['추가발주입력']
        for c in target_order:
            if c not in df_v5_v.columns: df_v5_v[c] = "" if c == "메모" else 0

        df_ed = df_v5_v[target_order].rename(columns=v_map)
        
        st.data_editor(
            df_ed, use_container_width=True, hide_index=True, key="v5_editor_final_ver",
            column_config={
                "상태": st.column_config.TextColumn("🚦 상태", width="small"),
                "추가발주": st.column_config.NumberColumn("➕ 추가발주", min_value=0),
                "권장발주": st.column_config.NumberColumn("💡 권장", disabled=True),
            }
        )
        
        if st.form_submit_button("✅ 1. 수량 및 메모 확정", use_container_width=True):
            edits = st.session_state.v5_editor_final_ver.get("edited_rows", {})
            for r_idx, val in edits.items():
                idx = df_v5_v.index[int(r_idx)]
                if "추가발주" in val: st.session_state.add_order_dict[idx] = int(val["추가발주"])
                if "메모" in val: st.session_state.df_raw.at[idx, "메모"] = str(val["메모"])
            st.rerun()

    # [5] 저장 및 다운로드
    c_save, c_down = st.columns(2)
    with c_save:
        if st.button("💾 2. 구글 시트 최종 저장", use_container_width=True, type="primary"):
            final_adds = st.session_state.add_order_dict
            if any(v > 0 for v in final_adds.values()):
                try:
                    sh = get_sheet()
                    ws_log = sh.worksheet("발주기록")
                    ws_hist = sh.worksheet("history")
                    now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M')
                    
                    rows = []
                    for idx, qty in final_adds.items():
                        if qty > 0:
                            rows.append([
                                now_s, str(df_v5.at[idx, item]), str(df_v5.at[idx, opt]), 
                                str(df_v5.at[idx, v_it]), int(df_v5.at[idx, avl]), 
                                int(df_v5.at[idx, "리오더잔량"]), int(qty), 
                                int(df_v5.at[idx, "발주권장"]), 
                                str(df_v5.at[idx, "메모"]).strip(), str(df_v5.at[idx, vnd])
                            ])
                    if rows:
                        ws_log.append_rows(rows)
                        ws_hist.append_rows(rows)
                        st.success(f"✅ {len(rows)}건 저장 완료!")
                        st.session_state.add_order_dict = {}
                        st.cache_data.clear()
                        time.sleep(1); st.rerun()
                except Exception as e: st.error(f"시트 저장 실패: {e}")
            else: st.warning("입력된 추가발주 수량이 없습니다.")

    with c_down:
        dl_list = [k for k, v in st.session_state.add_order_dict.items() if v > 0]
        if dl_list:
            df_dl = df_v5[df_v5.index.isin(dl_list)].copy()
            df_dl['수량'] = df_dl.index.map(st.session_state.add_order_dict)
            csv_file = df_dl[[vnd, item, opt, v_it, '수량']].to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button("📥 3. 발주서(CSV) 다운로드", csv_file, file_name=f"발주서_{datetime.now(KST).strftime('%m%d_%H%M')}.csv", use_container_width=True)
        else:
            st.button("📥 3. 다운로드 (내역없음)", disabled=True, use_container_width=True)



# ==================================================================
# [6단계: 추가 발주 히스토리 관리]
# ==================================================================
if st.session_state.get('analyzed'):
    st.divider()
    st.subheader("📜 6단계: 추가 발주 히스토리 관리")
    st.info("💡 5단계 화면 구성 그대로 조회하되, 메모 열만 제외하여 깔끔하게 보여줍니다.")

    h_f1, h_f2, h_f3 = st.columns([1.5, 1, 2])
    with h_f1:
        v6_date = st.date_input("🗓️ 조회 기간", value=(datetime.now(KST).date(), datetime.now(KST).date()), key="v6_date_pick")
    with h_f2:
        st.write(""); st.write("")
        v6_search_btn = st.button("🔍 내역 조회 실행", use_container_width=True, type="primary")

    if 'v6_storage' not in st.session_state: st.session_state.v6_storage = None

    if v6_search_btn:
        try:
            with st.spinner("📡 history 시트 로드 중..."):
                ws_h = get_sheet().worksheet("history")
                h_all = ws_h.get_all_values()
                if len(h_all) > 1:
                    df_h = pd.DataFrame(h_all[1:], columns=h_all[0])
                    df_h["date_only"] = df_h["발주시간"].astype(str).str.slice(0, 10)
                    start_s = v6_date[0].strftime('%Y-%m-%d')
                    end_s = v6_date[1].strftime('%Y-%m-%d') if len(v6_date) > 1 else start_s
                    
                    # 추가발주가 있는 내역만 필터링
                    q_col = "추가발주량" if "추가발주량" in df_h.columns else "추가발주"
                    df_h[q_col] = pd.to_numeric(df_h[q_col], errors='coerce').fillna(0).astype(int)
                    st.session_state.v6_storage = df_h[(df_h["date_only"] >= start_s) & (df_h["date_only"] <= end_s) & (df_h[q_col] > 0)].copy()
                else: st.session_state.v6_storage = None
        except Exception as e: st.error(f"조회 에러: {e}")

    if st.session_state.v6_storage is not None:
        df_h_view = st.session_state.v6_storage.copy()
        with h_f3:
            h_q = st.text_input("🔍 결과 내 검색 (상품/옵션)", key="v6_inner_search")
            if h_q: df_h_view = df_h_view[df_h_view["상품명"].str.contains(h_q, case=False) | df_h_view["옵션"].str.contains(h_q, case=False)]

        if not df_h_view.empty:
            # ⭐ 사장님 요청: 5단계 화면 구성 그대로 (메모 제외)
            # 순서: 발주시간, 업체명, 상품명, 옵션, 가용재고, 기존리오더, 추가발주
            h_target_q = "추가발주량" if "추가발주량" in df_h_view.columns else "추가발주"
            h_disp_cols = ["발주시간", "업체명", "상품명", "옵션", "가용재고", "기존리오더", h_target_q]
            
            # 숫자형 변환
            for num_c in ["가용재고", "기존리오더", h_target_q]:
                if num_c in df_h_view.columns:
                    df_h_view[num_c] = pd.to_numeric(df_h_view[num_c], errors='coerce').fillna(0).astype(int)

            st.dataframe(
                df_h_view[h_disp_cols].sort_values("발주시간", ascending=False),
                use_container_width=True, hide_index=True,
                column_config={
                    "가용재고": st.column_config.NumberColumn("📦 가용"),
                    "기존리오더": st.column_config.NumberColumn("🗒️ 기존"),
                    h_target_q: st.column_config.NumberColumn("➕ 추가발주")
                }
            )
        else: st.info("조회된 내역이 없습니다.")
        

# ------------------------------------------------------------------
# [7단계: 실시간 상황판] - 업데이트 버튼 클릭 시에만 갱신 (API 최적화)
# ------------------------------------------------------------------
import io
import pandas as pd
import streamlit as st

if st.session_state.get('analyzed'):
    st.divider()
    st.subheader("🚀 7단계: 실시간 리오더 최종 잔량 상황판")
    st.info("💡 [🔄 상황판 데이터 갱신] 버튼을 누를 때만 실시간 데이터를 집계합니다.")

    # [1] 데이터 로드 함수 (캐시 적용)
    @st.cache_data(ttl=600) # 10분간 캐시 유지 (버튼 클릭 시 clear_cache 함)
    def get_v7_minimal_memo():
        try:
            sh = get_sheet()
            # 발주기록
            ws_o = sh.worksheet("발주기록")
            o_all = ws_o.get_all_values()
            o_h_idx = 1 if len(o_all) > 1 and "상품명" in o_all[1] else 0
            df_o = pd.DataFrame(o_all[o_h_idx+1:], columns=o_all[o_h_idx]) if len(o_all) > 1 else pd.DataFrame()
            
            # 입고기록
            ws_r = sh.worksheet("입고기록")
            r_all = ws_r.get_all_values()
            r_h_idx = 1 if len(r_all) > 1 and "상품명" in r_all[1] else 0
            df_r = pd.DataFrame(r_all[r_h_idx+1:], columns=r_all[r_h_idx]) if len(r_all) > 1 else pd.DataFrame()
            
            return df_o, df_r
        except Exception as e:
            st.error(f"시트 연결 실패: {e}")
            return pd.DataFrame(), pd.DataFrame()

    # --- [2. 업데이트 제어 UI] ---
    c1, c2, c3, c4 = st.columns([1.5, 1, 1.5, 1.5])
    
    with c2:
        st.write(" ")
        # 🔄 이 버튼을 누를 때만 캐시를 비우고 새로 읽어옵니다.
        update_trigger = st.button("🔄 상황판 데이터 갱신", use_container_width=True, type="primary")
        if update_trigger:
            st.cache_data.clear() # 기존에 저장된 데이터 삭제
            st.rerun() # 새로고침하여 get_v7_minimal_memo() 재실행

    # 데이터 가져오기
    df_o_raw, df_r_raw = get_v7_minimal_memo()

    # 데이터가 비어있지 않을 때만 계산 시작
    if not df_o_raw.empty:
        q_col = next((c for c in ["추가발주", "추가발주량", "수량"] if c in df_o_raw.columns), df_o_raw.columns[6])
        date_col = next((c for c in ["날짜", "발주시간"] if c in df_o_raw.columns), df_o_raw.columns[0])

        # 데이터 전처리
        for df in [df_o_raw, df_r_raw]:
            if not df.empty:
                df[q_col] = pd.to_numeric(df[q_col], errors='coerce').fillna(0).astype(int)
                df['key'] = (df['상품명'].astype(str) + df['옵션'].astype(str)).str.replace(" ","").str.upper()
                df['short_date'] = pd.to_datetime(df[date_col], errors='coerce').dt.strftime('%m/%d').fillna('')

        # 1. 발주 집계 (이모지 제거 및 중복 제거)
        df_o_raw['memo_clean'] = df_o_raw.apply(
            lambda x: f"{x['short_date']} {x['메모']}" if str(x.get('메모','')).strip() else "", axis=1
        )
        df_orders = df_o_raw[df_o_raw[q_col] > 0].groupby(['key', '업체명', '상품명', '옵션'], as_index=False).agg({
            date_col: 'max', '공급처상품명': 'first', q_col: 'sum',
            'memo_clean': lambda x: " / ".join(dict.fromkeys(filter(None, x.astype(str))))
        }).rename(columns={q_col: '총리오더수량', 'memo_clean': '발주메모'})

        # 2. 입고 집계
        df_r_minus = df_r_raw[df_r_raw[q_col] < 0].copy() if not df_r_raw.empty else pd.DataFrame()
        if not df_r_minus.empty:
            df_r_minus['memo_clean'] = df_r_minus.apply(
                lambda x: f"{x['short_date']} {x[q_col]}{x['메모']}" if str(x.get('메모','')).strip() else f"{x['short_date']} {x[q_col]}개", axis=1
            )
            df_receives = df_r_minus.groupby('key').agg({
                q_col: lambda x: abs(x.sum()),
                'memo_clean': lambda x: " / ".join(dict.fromkeys(filter(None, x.astype(str))))
            }).reset_index().rename(columns={q_col: '입고수량', 'memo_clean': '입고메모'})
        else:
            df_receives = pd.DataFrame(columns=['key', '입고수량', '입고메모'])

        # 3. 통합
        df_total = pd.merge(df_orders, df_receives, on='key', how='left')
        df_total['입고수량'] = df_total['입고수량'].fillna(0).astype(int)
        df_total['미입고잔량'] = df_total['총리오더수량'] - df_total['입고수량']
        
        def combine_memos(row):
            m = []
            if row['발주메모']: m.append(f"[발주] {row['발주메모']}")
            if row['입고메모']: m.append(f"[입고] {row['입고메모']}")
            return " | ".join(m)
            
        df_total['메모이력'] = df_total.apply(combine_memos, axis=1)
        df_total = df_total[df_total['미입고잔량'] > 0].copy()

        # --- [4. 필터 UI (나머지 컬럼들)] ---
        with c1: search_date = st.date_input("📅 날짜 범위", value=[])
        with c3: search_prod = st.text_input("📦 상품 검색", placeholder="상품명/옵션")
        with c4:
            v_list = ["전체 업체"] + sorted(df_total["업체명"].unique().tolist())
            search_vendor = st.selectbox("🏭 업체 선택", v_list)

        # --- [5. 업체별 요약] ---
        st.write("#### 🏭 업체별 미입고 요약")
        v_summary = df_total.groupby("업체명")["미입고잔량"].sum().reset_index().sort_values("미입고잔량", ascending=False)
        if not v_summary.empty:
            m_cols = st.columns(5)
            for i, row in enumerate(v_summary.itertuples()):
                with m_cols[i % 5]:
                    st.metric(label=row.업체명, value=f"{int(row.미입고잔량):,}개")
        st.divider()

        # --- [6. 상세 리스트] ---
        df_disp = df_total.copy()
        if search_vendor != "전체 업체": df_disp = df_disp[df_disp["업체명"] == search_vendor]
        if search_prod: df_disp = df_disp[df_disp["상품명"].str.contains(search_prod, case=False) | df_disp["옵션"].str.contains(search_prod, case=False)]
        
        if not df_disp.empty:
            display_cols = ["날짜", "업체명", "상품명", "옵션", "공급처상품명", "총리오더수량", "입고수량", "미입고잔량", "메모이력"]
            st.dataframe(
                df_disp.sort_values("날짜", ascending=False).rename(columns={date_col: "날짜"}),
                use_container_width=True, hide_index=True,
                column_order=display_cols,
                column_config={
                    "총리오더수량": st.column_config.NumberColumn("총발주"),
                    "입고수량": st.column_config.NumberColumn("입고"),
                    "미입고잔량": st.column_config.NumberColumn("잔량"),
                    "메모이력": st.column_config.TextColumn("메모 (날짜/수량/내용)")
                }
            )
            
            # 엑셀 다운로드
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_disp[display_cols].to_excel(writer, index=False, sheet_name='미입고')
            st.download_button(label="📥 엑셀 다운로드", data=output.getvalue(), file_name="미입고현황.xlsx")
    else:
        st.info("💡 상황판을 갱신하려면 상단의 '상황판 데이터 갱신' 버튼을 눌러주세요.")
