import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
import time

# 1. 환경 설정 (KST 시간 고정)
KST = timezone(timedelta(hours=9))
now = datetime.now(KST)

st.set_page_config(layout="wide", page_title="저스트원 재고관리 v3.3")

# --- [공통 함수: 구글 시트 연동] ---
def get_sheet():
    try:
        from oauth2client.service_account import ServiceAccountCredentials
        import gspread
        creds_dict = dict(st.secrets["gcp_service_account"])
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        return client.open_by_key("1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg")
    except:
        return None

# 저장된 날짜 리스트 가져오기 (시각화용)
def get_saved_dates():
    sheet = get_sheet()
    if sheet:
        try:
            df_hist = pd.DataFrame(sheet.worksheet("발주기록").get_all_records())
            if not df_hist.empty and '날짜' in df_hist.columns:
                return df_hist['날짜'].unique().tolist()
        except: pass
    return []

# --- [세션 상태 초기화] ---
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'analyzed' not in st.session_state: st.session_state.analyzed = False
if 'p' not in st.session_state: st.session_state.p = {}
if 'add_order_dict' not in st.session_state: st.session_state.add_order_dict = {}

st.title("📦 저스트원 통합 재고 관리 v3.3")

tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 관리"])

with tab1:
    # --- 1~3단계: 데이터 업로드 및 분석 설정 ---
    st.subheader("📁 1~3단계: 데이터 업로드 및 분석 설정")
    up_file = st.file_uploader("엑셀/CSV 파일 업로드", type=['xlsx', 'xls', 'csv'], key="main_up")
    
    if st.button("🔄 시스템 전체 초기화", use_container_width=True):
        for key in list(st.session_state.keys()): del st.session_state[key]
        st.rerun()

    if up_file and st.session_state.df_raw is None:
        df = pd.read_csv(up_file) if up_file.name.endswith('.csv') else pd.read_excel(up_file)
        df.columns = df.columns.str.strip()
        st.session_state.df_raw = df
        st.rerun()

    if st.session_state.df_raw is not None:
        cols = st.session_state.df_raw.columns.tolist()
        def auto_idx(keys):
            for i, c in enumerate(cols):
                if any(k in str(c) for k in keys): return i
            return 0

        c1, c2, c3 = st.columns(3)
        with c1:
            so = st.selectbox("품절 여부", cols, index=auto_idx(['품절']))
            vn = st.selectbox("공급처", cols, index=auto_idx(['공급처']))
            vi = st.selectbox("공급처 상품명", cols, index=auto_idx(['공급처상품명']))
        with c2:
            it = st.selectbox("상품명", cols, index=auto_idx(['상품명']))
            op = st.selectbox("옵션", cols, index=auto_idx(['옵션']))
            stk = st.selectbox("정상재고", cols, index=auto_idx(['정상재고']))
        with c3:
            av = st.selectbox("가용재고", cols, index=auto_idx(['가용재고']))
            t3 = st.selectbox("3일 판매", cols, index=auto_idx(['3일']))
            t7 = st.selectbox("7일 판매", cols, index=auto_idx(['7일']))
        
        lt_val = st.number_input("⏳ 리드타임 (일)", value=7)
        ss_val = st.number_input("🛡️ 안전재고 (일)", value=3)

        if st.button("🚀 데이터 분석 시작", use_container_width=True, type="primary"):
            st.session_state.p = {'so': so, 'vn': vn, 'vi': vi, 'it': it, 'op': op, 'st': stk, 'av': av, 't3': t3, 't7': t7, 'lt': lt_val, 'ss': ss_val}
            st.session_state.analyzed = True
            st.rerun()

    # --- [4단계: 데이터 편집 및 재고 관리] ---
    if st.session_state.get('analyzed'):
        st.divider()
        st.subheader("📊 4단계: 데이터 편집 및 재고 관리")
        
        p = st.session_state.p
        df_work = st.session_state.df_raw.copy()
        
        # 숫자형 변환
        for c in [p['st'], p['av'], p['t7'], p['t3']]:
            df_work[c] = pd.to_numeric(df_work[c], errors='coerce').fillna(0).astype(int)
        
        if "리오더 수량" not in df_work.columns: df_work["리오더 수량"] = 0
        df_work["리오더 수량"] = pd.to_numeric(df_work["리오더 수량"], errors='coerce').fillna(0).astype(int)
        df_work["리오더 입고수량"] = 0

        # 일판매량 계산 (3일 판매 합계 기준)
        df_work['일판매량'] = (df_work[p['t3']] / 3).round(1)
        df_work['권장발주량'] = ((df_work['일판매량'] * (p['lt'] + p['ss'])) - (df_work[p['av']] + df_work['리오더 수량'])).clip(lower=0).astype(int)

        # 4단계 필터 및 기록 날짜 확인
        saved_dates = get_saved_dates()
        f4_c1, f4_c2, f4_c3 = st.columns([2, 1, 1])
        s4_q = f4_c1.text_input("🔍 4단계 상품명 검색", key="s4_search")
        m4_f = f4_c2.selectbox("4단계 상태 필터", ["전체보기", "정상만", "품절만"], index=1, key="m4_filter")
        
        with f4_c3:
            d4_h = st.date_input("🗓️ 입고 기록 날짜", now.date(), key="d4_date")
            if d4_h.strftime('%Y-%m-%d') in saved_dates:
                st.caption("✅ 선택한 날짜에 저장된 기록이 있습니다.")
            else:
                st.caption("⚪ 기록 없음")

        if m4_f == "정상만": df_work = df_work[~df_work[p['so']].astype(str).str.contains('품절', na=False)]
        elif m4_f == "품절만": df_work = df_work[df_work[p['so']].astype(str).str.contains('품절', na=False)]
        if s4_q: df_work = df_work[df_work[p['it']].astype(str).str.contains(s4_q, case=False) | df_work[p['op']].astype(str).str.contains(s4_q, case=False)]

        df_disp4 = df_work.rename(columns={p['so']: "품절", p['vn']: "공급쳐", p['vi']: "공급쳐 상품명", p['it']: "상품명", p['op']: "옵션", p['st']: "정상재고", p['av']: "가용재고", p['t3']: "3일 판매 합계"})
        cols4 = ["품절", "공급쳐", "상품명", "옵션", "공급쳐 상품명", "정상재고", "가용재고", "리오더 수량", "리오더 입고수량", "3일 판매 합계", "일판매량", "권장발주량"]

        with st.form("form_v4"):
            ed4 = st.data_editor(df_disp4[cols4], use_container_width=True, hide_index=True, key="ed4")
            if st.form_submit_button("💾 4단계 입고 반영 및 저장"):
                edits = st.session_state["ed4"].get("edited_rows", {})
                for r_idx, change in edits.items():
                    o_idx = df_work.index[int(r_idx)]
                    if "리오더 수량" in change: st.session_state.df_raw.at[o_idx, "리오더 수량"] = int(change["리오더 수량"])
                    if "리오더 입고수량" in change:
                        in_qty = int(change["리오더 입고수량"])
                        st.session_state.df_raw.at[o_idx, "리오더 수량"] = max(0, int(st.session_state.df_raw.at[o_idx, "리오더 수량"]) - in_qty)
                st.success("✅ 리오더 차감 완료!"); time.sleep(1); st.rerun()

        # ==========================================
        # --- [5단계: 최종 발주 및 기록 날짜 시각화] ---
        # ==========================================
        st.divider()
        st.subheader("📋 5단계: 최종 발주 리스트 요약")

        df_5 = st.session_state.df_raw.copy()
        # [원본 소스 데이터 매칭 유지] p 변수를 사용하여 데이터 로드
        for c in [p['av'], p['t3']]:
            df_5[c] = pd.to_numeric(df_5[c], errors='coerce').fillna(0).astype(int)

        # 5단계 필터 및 날짜 확인
        f5_c1, f5_c2, f5_c3 = st.columns([2, 1, 1])
        s5_q = f5_c1.text_input("🔍 5단계 상품명 검색", key="s5_search")
        m5_f = f5_c2.selectbox("5단계 상태 필터", ["전체보기", "정상만", "품절만"], index=1, key="m5_filter")
        
        with f5_c3:
            d5_h = st.date_input("🗓️ 발주 기록 날짜", now.date(), key="d5_date")
            if d5_h.strftime('%Y-%m-%d') in saved_dates:
                st.markdown(f"<span style='color:green; font-weight:bold;'>● {d5_h.strftime('%m/%d')} 기록 있음</span>", unsafe_allow_html=True)
            else:
                st.markdown(f"<span style='color:gray;'>○ 기록 없음</span>", unsafe_allow_html=True)

        if m5_f == "정상만": df_5 = df_5[~df_5[p['so']].astype(str).str.contains('품절', na=False)]
        elif m5_f == "품절만": df_5 = df_5[df_5[p['so']].astype(str).str.contains('품절', na=False)]
        if s5_q: df_5 = df_5[df_5[p['it']].astype(str).str.contains(s5_q, case=False) | df_5[p['op']].astype(str).str.contains(s5_q, case=False)]

        df_5['일판매량'] = (df_5[p['t3']] / 3).round(1)
        df_5['권장발주량'] = ((df_5['일판매량'] * (p['lt'] + p['ss'])) - (df_5[p['av']] + df_5['리오더 수량'])).clip(lower=0).astype(int)
        df_5['추가발주수량'] = df_5.index.map(st.session_state.add_order_dict).fillna(0).astype(int)

        def get_stat5(r):
            tot = r[p['av']] + r['리오더 수량']; day = r['일판매량']
            if day > 0:
                if tot < (day * 3): return "🚨 긴급"
                if tot < (day * 5): return "⚠️ 주의"
            return "✅ 정상"
        df_5['상태'] = df_5.apply(get_stat5, axis=1)

        # 화면 표시용 컬럼 재매칭 (사장님 원본 소스값 유지)
        df_disp5 = df_5.rename(columns={p['it']: "상품명", p['op']: "옵션", p['vi']: "공급쳐상품명", p['av']: "가용재고", "리오더 수량": "리오더수량", p['t3']: "3일 판매 합계"})
        cols5 = ["상태", "상품명", "옵션", "공급쳐상품명", "가용재고", "리오더수량", "3일 판매 합계", "일판매량", "추가발주수량", "권장발주량"]

        with st.form("form_v5"):
            ed5 = st.data_editor(df_disp5[cols5], use_container_width=True, hide_index=True, key="ed5")
            if st.form_submit_button("✅ 5단계 수량 확정 (리오더 합산)"):
                changes = st.session_state["ed5"].get("edited_rows", {})
                for r_idx, val in changes.items():
                    o_idx = df_5.index[int(r_idx)]
                    if "추가발주수량" in val:
                        add_v = int(val["추가발주수량"])
                        st.session_state.df_raw.at[o_idx, "리오더 수량"] += add_v
                        st.session_state.add_order_dict[o_idx] = add_v
                st.success("✅ 리오더 수량 업데이트 완료!"); time.sleep(1); st.rerun()

        if st.button("💾 구글 시트에 최종 발주 기록 저장", use_container_width=True):
            ready = df_5[(df_5['권장발주량'] + df_5['추가발주수량']) > 0]
            if not ready.empty:
                log_rows = [[d5_h.strftime('%Y-%m-%d'), r['상태'], r[p['it']], r[p['op']], r[p['vi']], int(r[p['av']]), int(r['리오더 수량']), int(r['추가발주수량']), int(r['권장발주량'])] for _, r in ready.iterrows()]
                sheet = get_sheet()
                if sheet: 
                    sheet.worksheet("발주기록").append_rows(log_rows)
                    st.success(f"✅ {d5_h.strftime('%m/%d')} 기록 저장 완료!"); time.sleep(1); st.rerun()

        # --- [6단계: 전체 히스토리 내역] ---
        st.divider()
        st.subheader("📜 6단계: 전체 히스토리 내역")
        if st.button("🔄 히스토리 불러오기", use_container_width=True):
            sheet = get_sheet()
            if sheet:
                df_hist = pd.DataFrame(sheet.worksheet("발주기록").get_all_records())
                st.dataframe(df_hist.sort_values(by=df_hist.columns[0], ascending=False), use_container_width=True, hide_index=True)
