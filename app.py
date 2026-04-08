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
# [6단계: 전체 히스토리 관리] - KeyError 방어 및 안전 조회 버전
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.subheader("📜 6단계: 추가 발주 히스토리 관리")
    st.info("💡 [히스토리] 시트에서 실제 '추가 발주'가 발생했던 내역만 불러옵니다.")

    f1, f2, f3, f4 = st.columns([1.2, 0.8, 1.5, 1.5])
    
    with f1:
        today = datetime.now(KST).date()
        d_range = st.date_input("🗓️ 1. 조회 범위", value=(today, today), key="v6_date_range")
    
    with f2:
        st.write(""); st.write("") 
        search_trigger = st.button("🔍 2. 내역 조회", use_container_width=True, type="primary")

    if 'v6_data' not in st.session_state: st.session_state.v6_data = None
    if 'v6_sessions' not in st.session_state: st.session_state.v6_sessions = []

    # [내역 조회 로직]
    if search_trigger:
        try:
            with st.spinner("📡 히스토리 데이터를 불러오는 중..."):
                worksheet = get_sheet().worksheet("히스토리")
                all_h = worksheet.get_all_values()
                
                if len(all_h) > 1:
                    df_all = pd.DataFrame(all_h[1:], columns=all_h[0])
                    
                    # 날짜 필터링
                    df_all["날짜_만"] = df_all["발주시간"].astype(str).str.slice(0, 10)
                    if isinstance(d_range, (list, tuple)) and len(d_range) == 2:
                        s_d, e_d = d_range[0].strftime('%Y-%m-%d'), d_range[1].strftime('%Y-%m-%d')
                    else:
                        s_d = e_d = d_range[0].strftime('%Y-%m-%d') if isinstance(d_range, (list, tuple)) else d_range.strftime('%Y-%m-%d')
                    
                    df_filtered = df_all[(df_all["날짜_만"] >= s_d) & (df_all["날짜_만"] <= e_d)].copy()
                    st.session_state.v6_data = df_filtered
                    st.session_state.v6_sessions = sorted(df_filtered["발주시간"].unique(), reverse=True)
                else:
                    st.session_state.v6_data = None
                    st.info("💡 해당 범위에 저장된 히스토리가 없습니다.")
        except Exception as e:
            st.error(f"📡 데이터를 불러오지 못했습니다: {e}")

    # [UI 및 데이터 표시]
    with f3: h_q = st.text_input("🔍 3. 상품명/옵션 검색", key="v6_search_q")
    with f4:
        if st.session_state.v6_sessions:
            session_options = ["📊 선택 범위 전체 합산"] + [f"{len(st.session_state.v6_sessions)-i}회차 ({t[5:16]})" for i, t in enumerate(st.session_state.v6_sessions)]
            sel_session_label = st.selectbox("📦 4. 회차 선택", session_options, key="v6_session_select")
        else:
            st.selectbox("📦 4. 회차 선택", ["조회 결과 없음"], disabled=True)
            sel_session_label = None

    if st.session_state.v6_data is not None and sel_session_label:
        df_display = st.session_state.v6_data.copy()
        
        # ⭐ [안전장치 1] 시트에 실제 존재하는 컬럼인지 먼저 확인 후 숫자 변환
        actual_cols = df_display.columns.tolist()
        num_targets = ["가용재고", "기존리오더", "추가발주량", "추가발주", "발주권장"]
        
        for col in num_targets:
            if col in actual_cols: # 컬럼이 있을 때만 변환해서 KeyError 방지
                df_display[col] = pd.to_numeric(df_display[col], errors='coerce').fillna(0).astype(int)

        # [회차별/합산별 데이터 정리]
        if sel_session_label == "📊 선택 범위 전체 합산":
            # 합산 시 사용할 안전한 agg 설정
            agg_logic = {"발주시간": "max", "메모": lambda x: " / ".join(set(filter(None, x.astype(str))))}
            # 수량 컬럼이 존재하면 합산 항목에 추가
            for c in ["추가발주량", "추가발주"]:
                if c in actual_cols: agg_logic[c] = "sum"
            for c in ["가용재고", "발주권장"]:
                if c in actual_cols: agg_logic[c] = "last"

            df_display = df_display.groupby(["업체명", "상품명", "옵션"], as_index=False).agg(agg_logic)
            display_title = "🗓️ 선택 범위 합계"
        else:
            target_time = st.session_state.v6_sessions[session_options.index(sel_session_label)-1]
            df_display = df_display[df_display["발주시간"] == target_time].copy()
            display_title = f"✅ {sel_session_label} 내역"

        # 검색 필터
        if h_q:
            df_display = df_display[df_display["상품명"].astype(str).str.contains(h_q, case=False) | 
                                    df_display["옵션"].astype(str).str.contains(h_q, case=False)]

        if not df_display.empty:
            st.write(f"#### {display_title}")
            
            # ⭐ [안전장치 2] 보여줄 때도 존재하는 컬럼만 선별해서 출력
            preferred_order = ["발주시간", "업체명", "상품명", "옵션", "추가발주량", "추가발주", "메모"]
            final_view_cols = [c for c in preferred_order if c in df_display.columns]
            
            st.dataframe(df_display[final_view_cols], use_container_width=True, hide_index=True)
            
            # CSV 다운로드
            csv_data = df_display[final_view_cols].to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button(
                label=f"📥 {display_title} CSV 다운로드", 
                data=csv_data, 
                file_name=f"발주히스토리_{datetime.now().strftime('%m%d')}.csv", 
                mime="text/csv",
                use_container_width=True
            )

        

# ------------------------------------------------------------------
# [7단계: 메모 컬럼 누락 복구 및 상세 필터 버전]
# ------------------------------------------------------------------
import io

if st.session_state.get('analyzed'):
    st.divider()
    st.subheader("🚀 7단계: 실시간 리오더 최종 잔량 상황판")

    @st.cache_data(ttl=600)
    def get_v7_memo_fixed_data():
        try:
            sh = get_sheet()
            ws_o = sh.worksheet("발주기록")
            o_all = ws_o.get_all_values()
            o_h_idx = 1 if len(o_all) > 1 and "상품명" in o_all[1] else 0
            df_o = pd.DataFrame(o_all[o_h_idx+1:], columns=o_all[o_h_idx])
            
            ws_r = sh.worksheet("입고기록")
            r_all = ws_r.get_all_values()
            r_h_idx = 1 if len(r_all) > 1 and "상품명" in r_all[1] else 0
            df_r = pd.DataFrame(r_all[r_h_idx+1:], columns=r_all[r_h_idx])
            
            return df_o, df_r
        except: return pd.DataFrame(), pd.DataFrame()

    # --- [상단 레이아웃] ---
    c1, c2, c3, c4 = st.columns([1.5, 0.8, 1.5, 1.5])
    with c1: search_date = st.date_input("📅 날짜 범위", value=[])
    with c2: 
        st.write(" ")
        if st.button("🔄 업데이트", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
    with c3: search_prod = st.text_input("📦 상품명 검색", placeholder="상품명/옵션 검색")
    with c4:
        df_o_raw, df_r_raw = get_v7_memo_fixed_data()
        vendor_list = ["전체 업체"] + (sorted(df_o_raw["업체명"].unique().tolist()) if not df_o_raw.empty else [])
        search_vendor = st.selectbox("🏭 업체 선택", vendor_list)

    if not df_o_raw.empty:
        # 컬럼명 유연하게 대처
        q_col = next((c for c in ["추가발주", "추가발주량", "수량"] if c in df_o_raw.columns), df_o_raw.columns[6])
        date_col = next((c for c in ["날짜", "발주시간"] if c in df_o_raw.columns), df_o_raw.columns[0])
        memo_col = "메모" if "메모" in df_o_raw.columns else None
        supply_col = "공급처상품명" if "공급처상품명" in df_o_raw.columns else None

        for df in [df_o_raw, df_r_raw]:
            df[q_col] = pd.to_numeric(df[q_col], errors='coerce').fillna(0).astype(int)
            df['key'] = (df['상품명'].astype(str) + df['옵션'].astype(str)).str.replace(" ","").str.upper()

        # [A] 발주 집계 (메모 데이터 보존 로직)
        agg_dict = {
            date_col: 'max',
            q_col: 'sum'
        }
        if supply_col: agg_dict[supply_col] = 'first'
        # 메모가 있으면 중복 제거해서 합치기
        if memo_col:
            agg_dict[memo_col] = lambda x: " / ".join(dict.fromkeys(filter(None, x.astype(str).str.strip())))

        df_orders = df_o_raw[df_o_raw[q_col] > 0].groupby(['key', '업체명', '상품명', '옵션'], as_index=False).agg(agg_dict)

        # [B] 입고 집계 (음수만)
        df_r_raw['actual_in'] = df_r_raw[q_col].apply(lambda x: abs(x) if x < 0 else 0)
        receive_sum = df_r_raw.groupby('key')['actual_in'].sum()

        # [C] 잔량 계산
        df_orders['미입고잔량'] = df_orders.apply(lambda x: x[q_col] - receive_sum.get(x['key'], 0), axis=1)
        df_final = df_orders[df_orders['미입고잔량'] > 0].copy()
        
        # --- [필터링] ---
        if search_vendor != "전체 업체":
            df_final = df_final[df_final["업체명"] == search_vendor]
        if search_prod:
            df_final = df_final[df_final["상품명"].str.contains(search_prod, case=False) | df_final["옵션"].str.contains(search_prod, case=False)]
        if len(search_date) == 2:
            df_final[date_col] = pd.to_datetime(df_final[date_col], errors='coerce')
            df_final = df_final[(df_final[date_col].dt.date >= search_date[0]) & (df_final[date_col].dt.date <= search_date[1])]

        # --- [최종 출력] ---
        if not df_final.empty:
            df_final = df_final.rename(columns={date_col: "날짜", q_col: "추가리오더"})
            # 실제 존재하는 컬럼만 표시 리스트에 추가
            actual_cols = ["날짜", "업체명", "상품명", "옵션"]
            if supply_col: actual_cols.append("공급처상품명")
            actual_cols.extend(["미입고잔량", "추가리오더"])
            if memo_col: actual_cols.append("메모")
            
            st.dataframe(
                df_final.sort_values("날짜", ascending=False),
                use_container_width=True,
                hide_index=True,
                column_order=actual_cols,
                column_config={
                    "미입고잔량": st.column_config.NumberColumn("🔢 잔량", format="%d"),
                    "추가리오더": st.column_config.NumberColumn("➕ 발주량", format="%d"),
                    "메모": st.column_config.TextColumn("📝 메모", width="large")
                }
            )
            
            # 엑셀 다운로드
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final[actual_cols].to_excel(writer, index=False, sheet_name='미입고현황')
            st.download_button(label="📥 엑셀 다운로드", data=output.getvalue(), file_name="미입고현황.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.info("내역이 없습니다.")
