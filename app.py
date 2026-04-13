import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import re
import unicodedata
import gspread
from datetime import datetime, timedelta, timezone

# 1. 환경 설정
KST = timezone(timedelta(hours=9))
st.set_page_config(layout="wide", page_title="저스트원 v8.2")

# [새로고침 방지]
components.html("<script>window.onbeforeunload = function() { return '변경사항이 저장되지 않을 수 있습니다.'; };</script>", height=0)

# --- [공통 함수] ---
def super_clean(t):
    if not t: return ""
    t = unicodedata.normalize('NFC', str(t))
    return re.sub(r'[^a-zA-Z0-9가-힣]', '', t).upper().strip()

def to_i(v):
    try: return int(float(str(v).replace(",", "").strip()))
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
        c_str = str(c).upper()
        if exclude_keys and any(k.upper() in c_str for k in exclude_keys): continue
        if any(k.upper() in c_str for k in keys): return i
    return 0

# --- [메인 화면 영역] ---
st.title("📦 저스트원 통합 재고 관리")

# 1단계: 업로드
st.header("1️⃣ 데이터 업로드")
if 'reset_trigger' not in st.session_state: st.session_state.reset_trigger = 0

col_up1, col_up2 = st.columns([8, 2])
with col_up1:
    up_file = st.file_uploader("📂 파일 업로드", type=['xlsx', 'xls', 'csv'], key=f"uploader_{st.session_state.reset_trigger}")
with col_up2:
    st.write("") # 간격 맞춤
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
                        key = super_clean(row[1]) + super_clean(row[2])
                        r_map[key] = r_map.get(key, 0) + (to_i(row[5]) + to_i(row[6]))
                st.session_state.r_map = r_map
        except Exception as e:
            st.error(f"⚠️ 파일 로드 실패: {e}")
            st.stop()

    cols = st.session_state.df_raw.columns.tolist()
    
    st.divider()
    
    # --- 2단계 & 3단계 영역 ---
    # 레이아웃을 5:5 비율로 설정
    col_step2, col_step3 = st.columns([1, 1])
    
    with col_step2:
        st.header("2️⃣ 필드 매핑")
        # 좌측/우측 5:5로 다시 세부 분리
        c_l, c_r = st.columns(2)
        with c_l:
            s_so = st.selectbox("🚫 품절 여부", cols, index=auto_idx(cols, ['품절']), key="so_box")
            s_vn = st.selectbox("🏭 공급처", cols, index=auto_idx(cols, ['공급처']), key="vn_box")
            s_vi = st.selectbox("🆔 공급처 상품명", cols, index=auto_idx(cols, ['공급처상품명']), key="vi_box")
            s_it = st.selectbox("📦 상품명", cols, index=auto_idx(cols, ['상품명']), key="it_box")
            s_op = st.selectbox("🎨 옵션", cols, index=auto_idx(cols, ['옵션']), key="op_box")
        with c_r:
            s_rd = st.selectbox("📅 등록일", cols, index=auto_idx(cols, ['등록일']), key="rd_box")
            s_st = st.selectbox("🏢 정상재고", cols, index=auto_idx(cols, ['정상재고']), key="st_box")
            s_av = st.selectbox("✅ 가용재고", cols, index=auto_idx(cols, ['가용재고']), key="av_box")
            s_t3 = st.selectbox("🔥 3일 발주합계", cols, index=auto_idx(cols, ['3일']), key="t3_box")
            s_t7 = st.selectbox("📅 7일 발주합계", cols, index=auto_idx(cols, ['7일', '1주']), key="t7_box")

    with col_step3:
        st.header("3️⃣ 수치 설정")
        lt = st.number_input("⏳ 리드타임 (입고 대기 기간)", value=7, key="lt_val")
        ss = st.number_input("🛡️ 안전재고 (최소 유지 재고)", value=3, key="ss_val")
        st.write("")
        st.write("")
        st.write("")
        analyze_btn = st.button("📊 분석 실행", type="primary", use_container_width=True)

    if analyze_btn:
        df = st.session_state.df_raw.copy()
        # 숫자 변환 (가용재고, 3일판매, 7일판매)
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

    # 4~5단계: 결과 편집 영역
    if 'analyzed_data' in st.session_state:
        st.divider()
        st.header("4️⃣~5️⃣ 발주 편집 및 저장")
        
        m = st.session_state.final_mapping
        
        f1, f2 = st.columns([3, 2])
        with f1:
            f_type = st.radio("🔍 필터", ["전체", "🚨 긴급(묶음)", "✅ 정상", "🚫 품절"], horizontal=True, key="f_radio")
        with f2:
            search_q = st.text_input("🔎 상품명 검색", key="s_input")

        df_view = st.session_state.analyzed_data.copy()
        if f_type == "🚨 긴급(묶음)":
            df_view = df_view[df_view['item_urgent_group'] == True]
        elif f_type == "✅ 정상":
            df_view = df_view[df_view[m['so']].astype(str).str.contains("정상", na=False)]
        elif f_type == "🚫 품절":
            df_view = df_view[df_view[m['so']].astype(str).str.contains("품절", na=False)]
        
        if search_q:
            df_view = df_view[df_view[m['it']].str.contains(search_q, case=False, na=False)]

        # '상태' 컬럼이 '공급처' 앞으로 오도록 설정
        t_cols = ['상태', m['vn'], m['it'], m['op'], m['vi'], m['av'], '리오더 수량', '입고차감', '추가발주', m['t3'], '1일 판매량', '권장발주수량', '메모']
        a_cols = [c for c in t_cols if c in df_view.columns]

        edited_df = st.data_editor(
            df_view[a_cols],
            use_container_width=True, hide_index=True, key="main_editor",
            column_config={
                "상태": st.column_config.TextColumn("📊 상태", width="small"),
                "입고차감": st.column_config.NumberColumn("📥 입고(-)", format="%d"),
                "추가발주": st.column_config.NumberColumn("➕ 추가", format="%d"),
            }
        )

        if st.button("💾 일괄 저장", type="primary", use_container_width=True):
            to_save = edited_df[(edited_df['입고차감'] != 0) | (edited_df['추가발주'] > 0)]
            if not to_save.empty:
                sh = get_sheet()
                ws_main = sh.worksheet("발주기록")
                now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M')
                rows = [[now_s, str(r[m['it']]), str(r[m['op']]), str(r[m['vi']]), 0, int(r['리오더 수량']), int(r['추가발주'])-int(r['입고차감']), int(r['권장발주수량']), str(r['메모']), str(r[m['vn']])] for _, r in to_save.iterrows()]
                ws_main.append_rows(rows)
                st.success("✅ 구글 시트에 저장되었습니다!")



st.divider()
        st.header("6️⃣ 저장 내역 검색 및 7️⃣ 데이터 내보내기")

        # 구글 시트에서 전체 기록 가져오기
        sh = get_sheet()
        if sh:
            ws_log = sh.worksheet("발주기록")
            raw_logs = ws_log.get_all_values()
            
            if len(raw_logs) > 1:
                df_logs = pd.DataFrame(raw_logs[1:], columns=raw_logs[0])
                
                # 검색 필터 영역
                search_col1, search_col2 = st.columns([1, 1])
                with search_col1:
                    q_item = st.text_input("🔎 검색 (상품명/옵션/공급처)", key="log_search_q")
                with search_col2:
                    # 날짜 검색 (오늘, 어제, 전체 등)
                    q_date = st.date_input("📅 날짜 선택", value=None, key="log_date_q")

                # 필터링 로직
                filtered_logs = df_logs.copy()
                if q_item:
                    filtered_logs = filtered_logs[
                        filtered_logs['상품명'].str.contains(q_item, case=False) | 
                        filtered_logs['옵션'].str.contains(q_item, case=False) |
                        filtered_logs['공급처'].str.contains(q_item, case=False)
                    ]
                if q_date:
                    target_date = q_date.strftime('%Y-%m-%d')
                    filtered_logs = filtered_logs[filtered_logs['날짜'].str.contains(target_date)]

                # 6단계: 검색 결과 표출
                st.subheader(f"📋 검색 결과 ({len(filtered_logs)}건)")
                st.dataframe(filtered_logs, use_container_width=True, hide_index=True)

                # 7단계: 내보내기 버튼
                st.subheader("📥 검색 결과 다운로드")
                import io
                csv_buffer = filtered_logs.to_csv(index=False).encode('utf-8-sig')
                
                st.download_button(
                    label="📂 검색된 내역 엑셀(CSV) 다운로드",
                    data=csv_buffer,
                    file_name=f"발주내역_검색결과_{datetime.now(KST).strftime('%m%d')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            else:
                st.info("💡 아직 저장된 발주 내역이 없습니다.")


# 7단계: 리오더 총수량 현황판 (6단계 아래에 추가)
        st.divider()
        st.header("7️⃣ 오늘의 리오더 발주 현황판")

        # 구글 시트에서 오늘 날짜 데이터만 추출
        today_str = datetime.now(KST).strftime('%Y-%m-%d')
        df_today = df_logs[df_logs['날짜'].str.contains(today_str)]

        if not df_today.empty:
            # 1. 상단 요약 수치 (Metrics)
            total_items = len(df_today)
            # 수량 컬럼을 숫자로 변환 (발주수량 컬럼명 확인 필요, 여기선 '발주수량' 가정)
            df_today['발주수량'] = df_today['발주수량'].apply(to_i)
            total_qty = df_today['발주수량'].sum()
            total_vendors = df_today['공급처'].nunique()

            m1, m2, m3 = st.columns(3)
            m1.metric("📦 총 발주 품목수", f"{total_items} 건")
            m2.metric("🔢 총 발주 총수량", f"{total_qty} 개")
            m3.metric("🏭 발주 공급처 수", f"{total_vendors} 곳")

            # 2. 공급처별 발주 현황 요약 표
            st.subheader("🏭 공급처별 발주 요약")
            vendor_summary = df_today.groupby('공급처').agg({
                '상품명': 'count',
                '발주수량': 'sum'
            }).rename(columns={'상품명': '품목 건수', '발주수량': '총 발주수량'}).reset_index()
            
            st.table(vendor_summary)

        else:
            st.warning(f"🔔 {today_str} 날짜로 저장된 발주 내역이 아직 없습니다.")
