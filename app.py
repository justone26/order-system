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

    # 4~5단계: 발주 편집 및 저장
    if 'analyzed_data' in st.session_state:
        st.divider()
        st.header("4️⃣~5️⃣ 발주 편집 및 저장")
        m = st.session_state.final_mapping
        
        f1, f2 = st.columns([3, 2])
        with f1:
            f_type = st.radio("🔍 필터", ["전체", "🚨 긴급(묶음)", "✅ 정상", "🚫 품절"], horizontal=True)
        with f2:
            search_q = st.text_input("🔎 상품명 검색")

        df_view = st.session_state.analyzed_data.copy()
        if f_type == "🚨 긴급(묶음)": df_view = df_view[df_view['item_urgent_group'] == True]
        elif f_type == "✅ 정상": df_view = df_view[df_view[m['so']].astype(str).str.contains("정상", na=False)]
        elif f_type == "🚫 품절": df_view = df_view[df_view[m['so']].astype(str).str.contains("품절", na=False)]
        if search_q: df_view = df_view[df_view[m['it']].str.contains(search_q, case=False, na=False)]

        t_cols = ['상태', m['vn'], m['it'], m['op'], m['vi'], m['av'], '리오더 수량', '입고차감', '추가발주', m['t3'], '1일 판매량', '권장발주수량', '메모']
        edited_df = st.data_editor(df_view[t_cols], use_container_width=True, hide_index=True, key="main_editor")

        if st.button("💾 일괄 저장", type="primary", use_container_width=True):
            to_save = edited_df[(edited_df['입고차감'] != 0) | (edited_df['추가발주'] > 0)]
            if not to_save.empty:
                sh = get_sheet()
                ws_main = sh.worksheet("발주기록")
                now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M')
                rows = [[now_s, str(r[m['it']]), str(r[m['op']]), str(r[m['vi']]), 0, int(r['리오더 수량']), int(r['추가발주'])-int(r['입고차감']), int(r['권장발주수량']), str(r['메모']), str(r[m['vn']])] for _, r in to_save.iterrows()]
                ws_main.append_rows(rows)
                st.success("✅ 구글 시트 저장 완료!")
                
# --- 6단계 & 7단계 시작 ---
        st.divider()
        sh = get_sheet()
        if sh:
            ws_log = sh.worksheet("발주기록")
            raw_logs = ws_log.get_all_values()
            
            if len(raw_logs) > 1:
                df_logs = pd.DataFrame(raw_logs[1:], columns=[c.strip() for c in raw_logs[0]])
                
                # [6단계: 검색 필터 영역]
                st.header("6️⃣ 저장 내역 상세 검색")
                s_col1, s_col2, s_col3 = st.columns([1, 1.5, 2.5])
                with s_col1:
                    q_date = st.date_input("📅 1차: 날짜 선택", value=None, key="search_date")
                with s_col2:
                    q_item_log = st.text_input("🔎 2차: 상품명 검색", key="search_item")
                with s_col3:
                    q_extra = st.text_input("📑 3차: 섹션 필터 (공급처/메모)", key="search_extra")

                # 필터링
                f_logs = df_logs.copy()
                d_col = next((c for c in f_logs.columns if '날짜' in c), f_logs.columns[0])
                i_col = next((c for c in f_logs.columns if '상품명' in c), None)
                v_col = next((c for c in f_logs.columns if '공급처' in c), None)
                m_col = next((c for c in f_logs.columns if '메모' in c), None)

                if q_date:
                    f_logs = f_logs[f_logs[d_col].str.contains(q_date.strftime('%Y-%m-%d'))]
                if q_item_log and i_col:
                    f_logs = f_logs[f_logs[i_col].str.contains(q_item_log, case=False)]
                if q_extra:
                    c_v = f_logs[v_col].str.contains(q_extra, case=False) if v_col else False
                    c_m = f_logs[m_col].str.contains(q_extra, case=False) if m_col else False
                    f_logs = f_logs[c_v | c_m]

                st.dataframe(f_logs.tail(30), use_container_width=True, hide_index=True)

                # [다운로드 버튼]
                if not f_logs.empty:
                    csv_data = f_logs.to_csv(index=False).encode('utf-8-sig')
                    st.download_button("📥 검색 결과 내역 다운로드 (CSV)", data=csv_data, file_name="발주기록_검색.csv", mime="text/csv", use_container_width=True)

                # [7단계: 현황판 영역 - 6단계 안쪽에 위치]
                st.divider()
                st.header("7️⃣ 오늘의 리오더 발주 현황판")
                today_str = datetime.now(KST).strftime('%Y-%m-%d')
                df_today = df_logs[df_logs[d_col].str.contains(today_str)]
                
                if not df_today.empty:
                    df_today = df_today.copy()
                    df_today['수량_num'] = df_today.iloc[:, 6].apply(to_i)
                    t_qty = df_today['수량_num'].sum()
                    t_vnd = df_today[v_col].nunique() if v_col else 0
                    
                    m1, m2 = st.columns(2)
                    m1.metric("오늘 총 발주수량", f"{t_qty} 개")
                    m2.metric("오늘 발주처 수", f"{t_vnd} 곳")
                    
                    if v_col:
                        v_sum = df_today.groupby(v_col)['수량_num'].sum().reset_index()
                        v_sum.columns = ['🏭 공급처', '📦 총 발주수량']
                        st.table(v_sum)
                else:
                    st.info("오늘 저장된 내역이 없습니다.")
            else:
                st.info("저장된 내역이 없습니다.")

        

                # --- 7단계 (현황판) 시작 (6단계 아래에 배치) ---
                st.divider()
                st.header("7️⃣ 오늘의 리오더 발주 현황판")
                today_str = datetime.now(KST).strftime('%Y-%m-%d')
                d_col = next((c for c in df_logs.columns if '날짜' in c), df_logs.columns[0])
                df_today = df_logs[df_logs[d_col].str.contains(today_str)]
                
                if not df_today.empty:
                    df_today = df_today.copy()
                    df_today['수량_num'] = df_today.iloc[:, 6].apply(to_i)
                    
                    t_qty = df_today['수량_num'].sum()
                    t_vnd = df_today[v_col].nunique() if v_col else 0
                    
                    m1, m2 = st.columns(2)
                    m1.metric("오늘 총 발주수량", f"{t_qty} 개")
                    m2.metric("오늘 발주처 수", f"{t_vnd} 곳")
                    
                    if v_col:
                        v_sum = df_today.groupby(v_col)['수량_num'].sum().reset_index()
                        v_sum.columns = ['🏭 공급처', '📦 총 발주수량']
                        st.table(v_sum)
                else:
                    st.info("오늘 저장된 내역이 아직 없습니다.")
            else:
                st.info("시트에 저장된 내역이 없습니다.")
