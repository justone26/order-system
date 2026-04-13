import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import re
import unicodedata
import gspread
from datetime import datetime, timedelta, timezone
import io

# 1. 환경 설정
KST = timezone(timedelta(hours=9))
st.set_page_config(layout="wide", page_title="저스트원 v9.2")

# [새로고침 방지]
components.html("<script>window.onbeforeunload = function() { return '변경사항이 저장되지 않을 수 있습니다.'; };</script>", height=0)

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

# --- [공통 데이터 로드] ---
# 여기서부터는 왼쪽 벽에 딱 붙여야 합니다! (공백 제거)
st.divider()
sh = get_sheet()

if sh:
    ws_log = sh.worksheet("발주기록")
    raw_logs = ws_log.get_all_values()
    
    # 데이터가 있는 경우에만 아래 단계들 실행
    if len(raw_logs) > 1:
        df_logs = pd.DataFrame(raw_logs[1:], columns=[c.strip() for c in raw_logs[0]])
        d_col = next((c for c in df_logs.columns if '날짜' in c), df_logs.columns[0])
        v_col = next((c for c in df_logs.columns if '공급처' in c), None)

        # 날짜/시간 전처리 (공통)
        df_logs['pure_dt'] = df_logs[d_col].str.strip()
        df_logs['pure_date'] = df_logs['pure_dt'].str.split(' ').str[0]
        df_logs['pure_time'] = df_logs['pure_dt'].str.split(' ').str[1].str[:5]


# --- [메인 화면] ---
st.title("📦 저스트원 통합 재고 관리")

# 1단계: 업로드
st.header("1️⃣ 데이터 업로드")
if 'reset_trigger' not in st.session_state: st.session_state.reset_trigger = 0

col_up1, col_up2 = st.columns([8, 2])
with col_up1:
    up_file = st.file_uploader("📂 파일 업로드", type=['xlsx', 'xls', 'csv'], key=f"uploader_{st.session_state.reset_trigger}")
with col_up2:
    st.write("") 
    st.write("") 
    if st.button("🔄 전체 초기화", use_container_width=True):
        for k in list(st.session_state.keys()): del st.session_state[k]
        st.rerun()

if up_file:
    if 'df_raw' not in st.session_state:
        try:
            df = pd.read_csv(up_file) if up_file.name.endswith('.csv') else pd.read_excel(up_file)
            df.columns = [str(c).strip() for c in df.columns]
            st.session_state.df_raw = df.fillna("")
            sh = get_sheet()
            if sh:
                ws = sh.worksheet("발주기록")
                all_vals = ws.get_all_values()
                r_map = {}
                if len(all_vals) > 1:
                    for row in all_vals[1:]:
                        if len(row) < 3: continue
                        key = super_clean(row[1]) + super_clean(row[2])
                        r_map[key] = r_map.get(key, 0) + (to_i(row[5]) + to_i(row[6]))
                st.session_state.r_map = r_map
        except Exception as e:
            st.error(f"⚠️ 파일 로드 실패: {e}")
            st.stop()

    cols = st.session_state.df_raw.columns.tolist()
    st.divider()
    
    # 2단계(매핑) & 3단계(설정) - 5:5 비율 유지
    col_step2, col_step3 = st.columns([1, 1])
    with col_step2:
        st.header("2️⃣ 필드 매핑")
        cl, cr = st.columns(2)
        with cl:
            s_so = st.selectbox("🚫 품절 여부", cols, index=auto_idx(cols, ['품절']), key="so_box")
            s_vn = st.selectbox("🏭 공급처", cols, index=auto_idx(cols, ['공급처']), key="vn_box")
            s_vi = st.selectbox("🆔 공급처 상품명", cols, index=auto_idx(cols, ['공급처상품명']), key="vi_box")
            s_it = st.selectbox("📦 상품명", cols, index=auto_idx(cols, ['상품명']), key="it_box")
            s_op = st.selectbox("🎨 옵션", cols, index=auto_idx(cols, ['옵션']), key="op_box")
        with cr:
            s_rd = st.selectbox("📅 등록일", cols, index=auto_idx(cols, ['등록일']), key="rd_box")
            s_st = st.selectbox("🏢 정상재고", cols, index=auto_idx(cols, ['정상재고']), key="st_box")
            s_av = st.selectbox("✅ 가용재고", cols, index=auto_idx(cols, ['가용재고']), key="av_box")
            s_t3 = st.selectbox("🔥 3일 발주합계", cols, index=auto_idx(cols, ['3일']), key="t3_box")
            s_t7 = st.selectbox("📅 7일 발주합계", cols, index=auto_idx(cols, ['7일', '1주']), key="t7_box")

    with col_step3:
        st.header("3️⃣ 수치 설정")
        lt = st.number_input("⏳ 리드타임 (입고 대기)", value=7, key="lt_val")
        ss = st.number_input("🛡️ 안전재고 (최소 유지)", value=3, key="ss_val")
        st.write("")
        if st.button("📊 분석 실행", type="primary", use_container_width=True):
            df = st.session_state.df_raw.copy()
            for c in [s_av, s_t3, s_t7]: df[c] = df[c].apply(to_i)
            df['clean_key'] = df.apply(lambda r: super_clean(r[s_it]) + super_clean(r[s_op]), axis=1)
            df['리오더 수량'] = df['clean_key'].map(st.session_state.r_map).fillna(0).astype(int)
            df['1일 판매량'] = df.apply(lambda r: int(round(r[s_t7]/7)) if r[s_t7]>0 else (int(round(r[s_t3]/3)) if r[s_t3]>0 else 0), axis=1)
            df['권장발주수량'] = ((df['1일 판매량'] * (lt + ss)) - (df[s_av] + df['리오더 수량'])).clip(lower=0).astype(int)
            df['상태'] = df['권장발주수량'].apply(lambda x: "🚨 긴급" if x > 0 else "✅ 정상")
            urgent_items = df[df['권장발주수량'] > 0][s_it].unique()
            df['item_urgent_group'] = df[s_it].isin(urgent_items)
            df['입고차감'] = 0
            df['추가발주'] = 0
            df['메모'] = ""
            st.session_state.analyzed_data = df.sort_values(by=['item_urgent_group', s_it, s_op], ascending=[False, True, True])
            st.session_state.final_mapping = {'vn':s_vn, 'it':s_it, 'op':s_op, 'vi':s_vi, 'av':s_av, 't3':s_t3, 'so':s_so}

# 4단계 분석 결과가 세션에 있을 때만 아래 단계들이 나타납니다.
    if 'analyzed_data' in st.session_state:
        df_res = st.session_state.analyzed_data
        m = st.session_state.final_mapping

        # ------------------------------------------------------------------
        # 4️⃣단계 & 5️⃣단계: 발주 편집 및 일괄 저장
        # ------------------------------------------------------------------
        st.divider()
        st.header("4️⃣~5️⃣ 발주 수량 편집 및 저장")
        
        # 편집용 데이터프레임 구성
        edited_df = st.data_editor(
            df_res,
            column_config={
                "상태": st.column_config.TextColumn("상태", width="small"),
                "추가발주": st.column_config.NumberColumn("➕ 추가발주", min_value=0, step=1),
                "입고차감": st.column_config.NumberColumn("➖ 입고차감", min_value=0, step=1),
                "메모": st.column_config.TextColumn("📝 메모", width="medium")
            },
            disabled=[c for c in df_res.columns if c not in ['추가발주', '입고차감', '메모']],
            hide_index=True,
            use_container_width=True,
            key="editor_v9"
        )

        if st.button("💾 일괄 저장 (시트 전송)", type="primary", use_container_width=True):
            to_save = edited_df[(edited_df['입고차감'] != 0) | (edited_df['추가발주'] > 0)]
            if not to_save.empty:
                try:
                    sh = get_sheet()
                    ws_main = sh.worksheet("발주기록")
                    now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M')
                    
                    rows = [[
                        now_s, 
                        str(r[m['it']]), 
                        str(r[m['op']]), 
                        str(r[m['vi']]), 
                        int(r[m['av']]), 
                        int(r['리오더 수량']), 
                        int(r['추가발주']) - int(r['입고차감']), # 실제 리오더 변동분
                        int(r['권장발주수량']), 
                        str(r['메모']), 
                        str(r[m['vn']])
                    ] for _, r in to_save.iterrows()]
                    
                    ws_main.append_rows(rows)
                    st.success("✅ 저장이 완료되었습니다! 수치 갱신을 위해 화면을 재로딩합니다.")
                    
                    # 수치 즉시 반영을 위한 세션 초기화 및 재실행
                    if 'r_map' in st.session_state: del st.session_state.r_map
                    st.rerun() 
                    
                except Exception as e:
                    st.error(f"❌ 저장 오류: {e}")
            else:
                st.warning("⚠️ 변경된 내역이 없습니다.")

        # ------------------------------------------------------------------
        # 6️⃣단계: 저장 내역 상세 검색 (사장님 요청 순서 및 필터 고정)
        # ------------------------------------------------------------------
        st.divider()
        st.header("6️⃣ 저장 내역 상세 검색")

        c1_6, c2_6 = st.columns([1, 1])
        with c1_6:
            q_date_6 = st.date_input("📅 1차: 날짜 선택", value=datetime.now(KST).date(), key="q_date_6")
        with c2_6:
            st.write("")
            btn_load_6 = st.button("🚀 데이터 조회하기", use_container_width=True, type="primary")

        if btn_load_6 or st.session_state.get('step6_active'):
            st.session_state.step6_active = True
            sh = get_sheet()
            if sh:
                ws = sh.worksheet("발주기록")
                raw = ws.get_all_values()
                if len(raw) > 1:
                    df_6 = pd.DataFrame(raw[1:], columns=[c.strip() for c in raw[0]])
                    df_6['pure_dt'] = df_6.iloc[:, 0].str.strip()
                    df_6['p_date'] = df_6['pure_dt'].str.split(' ').str[0]
                    df_6['p_time'] = df_6['pure_dt'].str.split(' ').str[1].str[:5]
                    
                    target_s = q_date_6.strftime('%Y-%m-%d')
                    res_date = df_6[df_6['p_date'] == target_s].copy()

                    if not res_date.empty:
                        sub1, sub2 = st.columns(2)
                        with sub1:
                            times = sorted(res_date['p_time'].dropna().unique(), reverse=True)
                            q_t = st.selectbox(f"⏰ 2차: 회차 선택 ({len(times)}회)", ["전체 보기"] + times)
                        with sub2:
                            q_i = st.text_input("🔎 3차: 상품명 검색")

                        display_6 = res_date.copy()
                        if q_t != "전체 보기": display_6 = display_6[display_6['p_time'] == q_t]
                        if q_i: display_6 = display_6[display_6.iloc[:, 1].str.contains(q_i, case=False)]

                        # [순서 고정] 날짜, 업체명, 상품명, 옵션, 공급처상품명, 가용, 기존, 입고(G), 추가, 권장, 메모
                        col_order = [
                            display_6.columns[0], display_6.columns[-1], "상품명", "옵션", 
                            "공급처상품명", "가용재고", "기존리오더", display_6.columns[6], 
                            "추가발주", "권장발주수량", "메모"
                        ]
                        final_cols = [c for c in col_order if c in display_6.columns]
                        st.dataframe(display_6[final_cols].iloc[::-1], use_container_width=True, hide_index=True)
                    else:
                        st.info("데이터가 없습니다.")

        # ------------------------------------------------------------------
        # 7️⃣단계: 실시간 리오더 최종 잔량 상황판 (미입고 전용)
        # ------------------------------------------------------------------
        st.divider()
        st.header("7️⃣ 실시간 리오더 최종 잔량 상황판")

        if st.button("📊 실시간 현황판 업데이트", key="btn_7_final", use_container_width=True):
            sh = get_sheet()
            if sh:
                ws = sh.worksheet("발주기록")
                raw = ws.get_all_values()
                if len(raw) > 1:
                    df_7 = pd.DataFrame(raw[1:], columns=[c.strip() for c in raw[0]])
                    df_7['qty'] = df_7.iloc[:, 6].apply(to_i)
                    v_col = df_7.columns[-1]
                    
                    # 업체/상품/옵션별 합산 후 미입고(>0)만 추출
                    df_rem = df_7.groupby([v_col, df_7.columns[1], df_7.columns[2]], as_index=False)['qty'].sum()
                    df_rem = df_rem[df_rem['qty'] > 0]
                    df_rem.columns = ['업체명', '상품명', '옵션', '잔량']

                    if not df_rem.empty:
                        v_sum = df_rem.groupby('업체명')['잔량'].sum().reset_index()
                        m_cols = st.columns(4)
                        for idx, r in enumerate(v_sum.itertuples()):
                            with m_cols[idx % 4]:
                                st.metric(r.업체명, f"{int(r.잔량)} 개")
                        st.dataframe(df_rem.sort_values('잔량', ascending=False), use_container_width=True, hide_index=True)
                    else:
                        st.success("✅ 모든 상품 입고 완료!")
