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
# [5단계: 최종 발주 요약] - 입고/발주/히스토리 3중 분산 저장 및 중복 방지
# ------------------------------------------------------------------
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📋 5단계: 최종 발주 리스트 요약")

    # [1] 실시간 잔량 로드 (발주기록 시트 기반 합산 데이터)
    reorder_map_v5, _ = get_realtime_data_v4(datetime.now(KST).date())
    df_v5 = st.session_state.df_raw.copy()

    # 상품 식별 키 생성
    def get_clean_key_v5(r):
        import unicodedata, re
        n = re.sub(r'[^a-zA-Z0-9가-힣]', '', unicodedata.normalize('NFC', str(r[item]))).upper().strip()
        o = re.sub(r'[^a-zA-Z0-9가-힣]', '', unicodedata.normalize('NFC', str(r[opt]))).upper().strip()
        return n + o

    df_v5['clean_key'] = df_v5.apply(get_clean_key_v5, axis=1)
    
    # 숫자 변환 및 품절 제외
    for col in [stk, avl, t3, t7]:
        df_v5[col] = pd.to_numeric(df_v5[col], errors='coerce').fillna(0).astype(int)
    df_v5 = df_v5[~df_v5[s_out].astype(str).str.contains('품절', na=False)]
    
    if '메모' not in df_v5.columns: df_v5['메모'] = ""
    if 'add_order_dict' not in st.session_state: st.session_state.add_order_dict = {}

    # [2] 화면 표시용 계산 (기존잔량 = 발주기록 탭 수치)
    df_v5["기존잔량"] = df_v5['clean_key'].map(reorder_map_v5).fillna(0).astype(int)
    df_v5['추가발주입력'] = df_v5.index.map(st.session_state.add_order_dict).fillna(0).astype(int)
    df_v5['총합계'] = df_v5["기존잔량"] + df_v5['추가발주입력']

    # [3] 필터 및 검색
    f1, f2 = st.columns([1, 2])
    with f1: m5_f = st.selectbox("🚦 상태 필터", ["✅ 전체보기", "🚨 추가발주 입력됨"], key="v5_f")
    with f2: s5_q = st.text_input("🔍 검색", key="v5_s")

    df_v5_v = df_v5.copy()
    if s5_q:
        df_v5_v = df_v5_v[df_v5_v[item].astype(str).str.contains(s5_q, case=False) | df_v5_v[opt].astype(str).str.contains(s5_q, case=False)]
    if m5_f == "🚨 추가발주 입력됨": 
        df_v5_v = df_v5_v[df_v5_v['추가발주입력'] > 0]
    
    # --- 에디터 폼 시작 ---
    with st.form("v5_optimized_form"):
        v_map = {item: "상품명", opt: "옵션", avl: "가용재고", "기존잔량": "기존잔량", "추가발주입력": "추가발주", "총합계": "총합계", "메모": "메모"}
        actual_cols = [item, opt, avl, "기존잔량", "추가발주입력", "총합계", "메모"]
        
        df_ed = df_v5_v[actual_cols].rename(columns=v_map)
        st.data_editor(
            df_ed, use_container_width=True, hide_index=True, key="v5_editor_final",
            column_config={
                "추가발주": st.column_config.NumberColumn("➕ 추가발주", min_value=0),
                "기존잔량": st.column_config.NumberColumn("📦 기존잔량", disabled=True),
                "총합계": st.column_config.NumberColumn("📊 총합계", disabled=True),
            }
        )
        
        confirm_btn = st.form_submit_button("✅ 1. 수량 확정 및 합계 확인", use_container_width=True)
        if confirm_btn:
            edits = st.session_state.v5_editor_final.get("edited_rows", {})
            for r_idx, val in edits.items():
                idx = df_v5_v.index[int(r_idx)]
                if "추가발주" in val: st.session_state.add_order_dict[idx] = int(val["추가발주"])
                if "메모" in val: st.session_state.df_raw.at[idx, "메모"] = str(val["메모"])
            st.rerun()

    # [4] 분산 저장 및 다운로드
    c_save, c_down = st.columns(2)
    
    with c_save:
        if st.button("💾 2. 구글 시트 분산 저장 (발주/히스토리)", use_container_width=True, type="primary"):
            # 세션에 저장된 추가 발주 수량 확인
            final_add_dict = st.session_state.add_order_dict
            
            if any(v > 0 for v in final_add_dict.values()):
                try:
                    sh = get_sheet()
                    ws_order = sh.worksheet("발주기록")
                    ws_hist = sh.worksheet("히스토리")
                    
                    now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M')
                    order_rows = []    # 발주기록용 (총합계 저장)
                    history_rows = []  # 히스토리용 (추가분만 저장)
                    
                    for idx, add_qty in final_add_dict.items():
                        if add_qty > 0:
                            total_qty = int(df_v5.at[idx, "기존잔량"]) + int(add_qty)
                            
                            # 공통 데이터 추출
                            row_base = [
                                now_s, 
                                str(df_v5.at[idx, item]), 
                                str(df_v5.at[idx, opt]), 
                                str(df_v5.at[idx, v_it]), 
                                int(df_v5.at[idx, avl]), 
                                0 # F열(기존리오더)은 0으로 고정하여 중복방지
                            ]
                            memo_val = str(df_v5.at[idx, "메모"])
                            vnd_val = str(df_v5.at[idx, vnd])

                            # A. 발주기록용 (총 리오더 수량 저장)
                            order_rows.append(row_base + [total_qty, 0, memo_val, vnd_val])
                            
                            # B. 히스토리용 (이번에 추가한 수량만 저장)
                            history_rows.append(row_base + [int(add_qty), 0, memo_val, vnd_val])
                    
                    if order_rows:
                        # 시트 저장 (추가된 상품만!)
                        ws_order.append_rows(order_rows)
                        ws_hist.append_rows(history_rows)
                        
                        st.success(f"✅ {len(order_rows)}건 분산 저장 완료! (발주기록/히스토리)")
                        st.cache_data.clear()
                        st.session_state.add_order_dict = {} 
                        time.sleep(1); st.rerun()
                except Exception as e:
                    st.error(f"시트 저장 실패: {e}")
            else:
                st.warning("저장할 추가 발주 수량이 없습니다. 1번 확정을 먼저 눌러주세요.")

    with c_down:
        # 다운로드 (추가 수량이 입력된 상품들만)
        download_list = [k for k, v in st.session_state.add_order_dict.items() if v > 0]
        if download_list:
            df_down = df_v5[df_v5.index.isin(download_list)].copy()
            df_down['최종수량'] = df_down.index.map(st.session_state.add_order_dict)
            
            csv_data = df_down[[vnd, item, opt, v_it, '최종수량']].rename(columns={
                vnd: "공급처", item: "상품명", opt: "옵션", v_it: "공급처상품명", "최종수량": "수량"
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
# [6단계: 전체 히스토리 관리] - [히스토리] 시트 전용 조회 버전
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.subheader("📜 6단계: 추가 발주 히스토리 관리")
    st.info("💡 이곳은 '추가 발주'가 발생하여 [히스토리] 시트에 저장된 내역만 표시합니다.")

    f1, f2, f3, f4 = st.columns([1.2, 0.8, 1.5, 1.5])
    
    with f1:
        today = datetime.now(KST).date()
        d_range = st.date_input("🗓️ 1. 조회 범위", value=(today, today), key="v6_date_range")
    
    with f2:
        st.write(""); st.write("") 
        search_trigger = st.button("🔍 2. 내역 조회", use_container_width=True, type="primary")

    if 'v6_data' not in st.session_state: st.session_state.v6_data = None
    if 'v6_sessions' not in st.session_state: st.session_state.v6_sessions = []
    if 'v6_display_text' not in st.session_state: st.session_state.v6_display_text = ""

    # [내역 조회 로직]
    if search_trigger:
        try:
            with st.spinner("📡 히스토리 데이터를 불러오는 중..."):
                # ⭐ 핵심변경: "발주기록"이 아닌 "히스토리" 시트를 읽어옵니다.
                worksheet = get_sheet().worksheet("히스토리")
                all_h = worksheet.get_all_values()
                
                if len(all_h) > 1:
                    df_all = pd.DataFrame(all_h[1:])
                    
                    # 5단계 저장 컬럼 순서 (10개): 
                    # [0]발주시간, [1]상품명, [2]옵션, [3]공급처상품명, [4]가용재고, [5]기존리오더(0), [6]추가발주량, [7]발주권장(0), [8]메모, [9]업체명
                    df_all.columns = [
                        "발주시간", "상품명", "옵션", "공급처상품명", 
                        "가용재고", "기존리오더", "추가발주량", "발주권장", "메모", "업체명"
                    ]
                    
                    df_all["날짜_만"] = df_all["발주시간"].astype(str).str.slice(0, 10)
                    curr_str = datetime.now(KST).strftime('%Y-%m-%d')
                    
                    # 날짜 필터링 처리
                    if isinstance(d_range, (list, tuple)) and len(d_range) == 2:
                        s_d, e_d = d_range[0].strftime('%Y-%m-%d'), d_range[1].strftime('%Y-%m-%d')
                        st.session_state.v6_display_text = f"🗓️ {s_d} ~ {e_d}"
                    elif isinstance(d_range, (list, tuple)) and len(d_range) == 1:
                        s_d, e_d = d_range[0].strftime('%Y-%m-%d'), curr_str
                        st.session_state.v6_display_text = f"🗓️ {s_d} ~ {e_d} (오늘까지)"
                    else:
                        s_d, e_d = "0000-00-00", "9999-99-99"
                        st.session_state.v6_display_text = "🗓️ 전체 내역"
                    
                    df_filtered = df_all[(df_all["날짜_만"] >= s_d) & (df_all["날짜_만"] <= e_d)].copy()
                    st.session_state.v6_data = df_filtered
                    st.session_state.v6_sessions = sorted(df_filtered["발주시간"].unique(), reverse=True)
                else:
                    st.session_state.v6_data = None
                    st.info("💡 히스토리 시트에 저장된 내역이 없습니다.")
        except Exception as e:
            st.error(f"📡 데이터를 불러오지 못했습니다: {e}")

    # [검색 및 회차 선택 UI]
    with f3: h_q = st.text_input("🔍 3. 상품명/옵션 검색", key="v6_search_q")
    with f4:
        if st.session_state.v6_sessions:
            session_options = ["📊 선택 범위 전체 합산"] + [f"{len(st.session_state.v6_sessions)-i}회차 ({t[5:16]})" for i, t in enumerate(st.session_state.v6_sessions)]
            sel_session_label = st.selectbox("📦 4. 회차 선택", session_options, key="v6_session_select")
        else:
            st.selectbox("📦 4. 회차 선택", ["조회 결과 없음"], disabled=True)
            sel_session_label = None

    # [데이터 표시]
    if st.session_state.v6_data is not None and sel_session_label:
        df_display = st.session_state.v6_data.copy()
        
        # 숫자 타입 변환
        num_cols = ["가용재고", "기존리오더", "추가발주량", "발주권장"]
        for col in num_cols:
            df_display[col] = pd.to_numeric(df_display[col], errors='coerce').fillna(0).astype(int)

        if sel_session_label == "📊 선택 범위 전체 합산":
            display_title = st.session_state.v6_display_text + " 추가발주 합계"
            df_display = df_display.groupby(["업체명", "상품명", "옵션", "공급처상품명"], as_index=False).agg({
                "발주시간": "max", 
                "가용재고": "last",
                "기존리오더": "last",
                "추가발주량": "sum", # ⭐ 합산 시 추가발주량을 더함
                "발주권장": "last",
                "메모": lambda x: " / ".join(set(filter(None, x.astype(str))))
            })
        else:
            target_time = st.session_state.v6_sessions[session_options.index(sel_session_label)-1]
            df_display = df_display[df_display["발주시간"] == target_time].copy()
            display_title = f"✅ {sel_session_label} 상세 내역"

        # 검색어 필터
        if h_q:
            df_display = df_display[
                df_display["상품명"].astype(str).str.contains(h_q, case=False) | 
                df_display["옵션"].astype(str).str.contains(h_q, case=False)
            ]

        if not df_display.empty:
            st.write(f"#### {display_title}")
            # 표시 순서 최적화
            view_order = ["발주시간", "업체명", "상품명", "옵션", "추가발주량", "메모", "공급처상품명"]
            st.dataframe(df_display[view_order], use_container_width=True, hide_index=True)
            
            csv_data = df_display[view_order].to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button(
                label=f"📥 {display_title} CSV 다운로드", 
                data=csv_data, 
                file_name=f"추가발주히스토리_{datetime.now().strftime('%m%d')}.csv", 
                mime="text/csv",
                use_container_width=True
            )



        

# ------------------------------------------------------------------
# [7단계: 실시간 리오더 최종 잔량 상황판] - 발주/입고 통합 계산 버전
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.subheader("🚀 7단계: 실시간 리오더 최종 잔량 상황판")

    # [1] 데이터 로드 함수 (발주기록 & 입고기록 두 군데 읽기)
    @st.cache_data(ttl=600) # 10분간 캐시
    def get_integrated_data():
        try:
            sh = get_sheet()
            # A. 발주기록 로드 (누적 발주량)
            ws_order = sh.worksheet("발주기록")
            order_data = ws_order.get_all_values()
            df_o = pd.DataFrame(order_data[1:], columns=order_data[0]) if len(order_data) > 1 else pd.DataFrame()
            
            # B. 입고기록 로드 (누적 입고량)
            ws_receive = sh.worksheet("입고기록")
            receive_data = ws_receive.get_all_values()
            df_r = pd.DataFrame(receive_data[1:], columns=receive_data[0]) if len(receive_data) > 1 else pd.DataFrame()
            
            return df_o, df_r
        except Exception as e:
            st.error(f"데이터 로드 실패: {e}")
            return pd.DataFrame(), pd.DataFrame()

    df_o, df_r = get_integrated_data()

    if not df_o.empty:
        # 1. 수치 데이터 정제 (발주기록)
        # 5단계에서 G열에 저장한 '추가발주' 컬럼이 실제 발주 수량입니다.
        df_o["추가발주"] = pd.to_numeric(df_o["추가발주"], errors='coerce').fillna(0).astype(int)
        df_o['key'] = (df_o['상품명'] + df_o['옵션']).str.replace(" ","").str.upper()
        
        # 2. 입고 데이터 정제 (입고기록)
        if not df_r.empty:
            # 입고기록 시트의 컬럼명에 맞춰 '입고수량'을 숫자로 변환
            # (입고기록 시트의 컬럼명이 '입고수량'인지 확인 필요, 아닐 경우 수정)
            rcv_col = "입고수량" if "입고수량" in df_r.columns else df_r.columns[6] 
            df_r[rcv_col] = pd.to_numeric(df_r[rcv_col], errors='coerce').fillna(0).astype(int)
            df_r['key'] = (df_r['상품명'] + df_r['옵션']).str.replace(" ","").str.upper()
            
            # 입고 합계 계산
            receive_sum = df_r.groupby('key')[rcv_col].sum()
        else:
            receive_sum = pd.Series()

        # 3. 발주 합계 계산 (상품별 총 발주량)
        order_summary = df_o.groupby(['key', '업체명', '상품명', '옵션', '공급처상품명'], as_index=False).agg({
            '발주시간': 'max',
            '추가발주': 'sum',
            '메모': lambda x: " / ".join(dict.fromkeys(filter(None, x.astype(str))))
        })

        # 4. ⭐ 최종 잔량 계산: (발주합계 - 입고합계)
        def calc_remain(row):
            total_ordered = row['추가발주']
            total_received = receive_sum.get(row['key'], 0)
            return max(0, total_ordered - total_received) # 마이너스 방지

        order_summary['최종잔량'] = order_summary.apply(calc_remain, axis=1)

        # [5] 필터 UI
        f1, f2, f3 = st.columns([1.5, 1.5, 1])
        with f1: q_v7 = st.text_input("🔍 상품명/옵션 검색", key="v7_qs")
        with f2: v_choice = st.selectbox("🏭 업체 필터", ["전체 업체"] + sorted(order_summary["업체명"].unique().tolist()), key="v7_vs")
        with f3: 
            st.write(""); st.write("")
            if st.button("🔄 데이터 새로고침", use_container_width=True, type="primary"):
                st.cache_data.clear()
                st.rerun()

        # 필터 적용
        df_f = order_summary.copy()
        if v_choice != "전체 업체": df_f = df_f[df_f["업체명"] == v_choice]
        if q_v7: df_f = df_f[df_f["상품명"].str.contains(q_v7, case=False) | df_f["옵션"].str.contains(q_v7, case=False)]
        
        # 잔량이 있는 것만 보기 (선택 사항)
        show_all = st.checkbox("✅ 입고 완료(잔량 0) 상품도 표시", value=False)
        if not show_all:
            df_f = df_f[df_f["최종잔량"] > 0]

        if not df_f.empty:
            # [6] 업체별 대시보드 (메트릭)
            st.write("### 📊 업체별 미입고 현황")
            df_v_sum = df_f.groupby("업체명")["최종잔량"].sum().reset_index().sort_values("최종잔량", ascending=False)
            
            v_cols = st.columns(4)
            for i, r in enumerate(df_v_sum.itertuples()):
                with v_cols[i % 4]:
                    st.metric(label=r.업체명, value=f"{int(r.최종잔량):,} 개")

            # [7] 상세 리스트 출력
            st.write("#### 📋 미입고 상세 리스트")
            st.dataframe(
                df_f.sort_values("발주시간", ascending=False),
                use_container_width=True,
                hide_index=True,
                column_order=["발주시간", "업체명", "상품명", "옵션", "추가발주", "최종잔량", "메모"],
                column_config={
                    "발주시간": st.column_config.TextColumn("🕒 최종발주일"),
                    "추가발주": st.column_config.NumberColumn("📦 총 발주량"),
                    "최종잔량": st.column_config.NumberColumn("🔢 미입고 잔량", format="%d", help="총 발주량에서 입고기록의 수량을 뺀 값입니다."),
                    "메모": st.column_config.TextColumn("📝 비고", width="medium")
                }
            )
        else:
            st.info("조회된 미입고 내역이 없습니다.")
    else:
        st.warning("발주 기록이 없습니다. 5단계에서 저장을 먼저 진행해주세요.")
