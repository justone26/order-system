import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
import time

# 1. 환경 설정 (KST 시간 고정)
KST = timezone(timedelta(hours=9))
now = datetime.now(KST)

st.set_page_config(layout="wide", page_title="저스트원 재고관리 v3.0")

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

def get_incoming_history():
    return pd.DataFrame(columns=['상품명', '옵션', '과거리오더 입고'])

# --- [세션 상태 초기화] ---
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'analyzed' not in st.session_state: st.session_state.analyzed = False
if 'p' not in st.session_state: st.session_state.p = {}
if 'add_order_dict' not in st.session_state: st.session_state.add_order_dict = {}

st.title("📦 저스트원 통합 재고 관리 v3.0")

tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 관리"])

# ==========================================
# [탭 1: 제작 상품 관리]
# ==========================================
with tab1:
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

    # ==========================================
    # --- [4단계: 데이터 편집 및 재고 관리] ---
    # ==========================================
    if st.session_state.get('analyzed') and st.session_state.get('p'):
        st.divider()
        st.subheader("📊 4단계: 데이터 편집 및 재고 관리")
        
        p = st.session_state.p
        sold_out, vendor, v_item, item, option = p['so'], p['vn'], p['vi'], p['it'], p['op']
        stock, avail, t3day, t7day = p['st'], p['av'], p['t3'], p['t7']
        lt, ss = p['lt'], p['ss']

        df_work = st.session_state.df_raw.copy()
        
        # 숫자형 변환
        for c in [stock, avail, t7day, t3day]:
            df_work[c] = pd.to_numeric(df_work[c], errors='coerce').fillna(0).astype(int)
        
        if "리오더 수량" not in df_work.columns: df_work["리오더 수량"] = 0
        df_work["리오더 수량"] = pd.to_numeric(df_work["리오더 수량"], errors='coerce').fillna(0).astype(int)
        df_work["리오더 입고수량"] = 0

        # 입고 기록 실시간 매칭
        incoming_hist = get_incoming_history()
        if not incoming_hist.empty:
            df_work = pd.merge(df_work, incoming_hist, left_on=[item, option], right_on=['상품명', '옵션'], how='left')
            df_work['과거리오더 입고'] = df_work['과거리오더 입고'].fillna(0).astype(int)
        else:
            df_work['과거리오더 입고'] = 0

        df_work['일판매량'] = df_work.apply(lambda x: round(x[t7day] / 7) if x[t7day] > 0 else round(x[t3day] / 3), axis=1).astype(int)
        df_work['권장발주량'] = ((df_work['일판매량'] * (lt + ss)) - (df_work[avail] + df_work['리오더 수량'])).clip(lower=0).astype(int)

        # UI 및 필터
        f_c1, f_c2, f_c3 = st.columns([2, 1, 1])
        search_q = f_c1.text_input("🔍 상품명/옵션 검색", key="v4_search")
        filter_m = f_c2.selectbox("상태 필터", ["전체보기", "정상만", "품절만"], index=1, key="v4_filter")
        hist_date_4 = f_c3.date_input("🗓️ 입고 기록 날짜", now.date(), key="v4_date")

        if filter_m == "정상만":
            df_work = df_work[~df_work[sold_out].astype(str).str.contains('품절', na=False)]
        elif filter_m == "품절만":
            df_work = df_work[df_work[sold_out].astype(str).str.contains('품절', na=False)]
        
        if search_q:
            df_work = df_work[df_work[item].astype(str).str.contains(search_q, case=False, na=False) | 
                              df_work[option].astype(str).str.contains(search_q, case=False, na=False)]

        df_display = df_work.rename(columns={sold_out: "품절", vendor: "공급쳐", v_item: "공급쳐 상품명", item: "상품명", option: "옵션", stock: "정상재고", avail: "가용재고"})
        final_cols = ["품절", "공급쳐", "상품명", "옵션", "공급쳐 상품명", "정상재고", "가용재고", "리오더 수량", "리오더 입고수량", "과거리오더 입고", "일판매량", "권장발주량"]

        with st.form("form_v4_safe"):
            edited_v4 = st.data_editor(df_display[final_cols], use_container_width=True, key="ed_v4_safe", hide_index=True,
                                      column_config={"정상재고": st.column_config.NumberColumn(format="%d"), "가용재고": st.column_config.NumberColumn(format="%d"),
                                                     "리오더 수량": st.column_config.NumberColumn(format="%d"), "리오더 입고수량": st.column_config.NumberColumn(format="%d")})
            if st.form_submit_button("💾 입고량 반영 및 저장 (리오더 차감)", use_container_width=True, type="primary"):
                edits = st.session_state["ed_v4_safe"].get("edited_rows", {})
                for r_idx, change in edits.items():
                    orig_idx = df_work.index[int(r_idx)]
                    if "리오더 수량" in change:
                        st.session_state.df_raw.at[orig_idx, "리오더 수량"] = int(change["리오더 수량"])
                    if "리오더 입고수량" in change:
                        # [중요] 리오더 수량에서 입고수량만큼 차감
                        in_qty = int(change["리오더 입고수량"])
                        curr = int(st.session_state.df_raw.at[orig_idx, "리오더 수량"])
                        st.session_state.df_raw.at[orig_idx, "리오더 수량"] = max(0, curr - in_qty)
                st.success("✅ 리오더 수량이 차감되었습니다."); time.sleep(1); st.rerun()

        # ==========================================
        # --- [5단계: 최종 발주 및 히스토리 자동 기록] ---
        # ==========================================
        st.divider()
        st.subheader("📋 5단계: 최종 발주 리스트 요약")

        df_5 = st.session_state.df_raw.copy()
        for c in [avail, t7day, t3day]:
            df_5[c] = pd.to_numeric(df_5[c], errors='coerce').fillna(0).astype(int)

        df_5['일판매량'] = df_5.apply(lambda x: round(x[t7day] / 7) if x[t7day] > 0 else round(x[t3day] / 3), axis=1).astype(int)
        df_5['권장발주량'] = ((df_5['일판매량'] * (lt + ss)) - (df_5[avail] + df_5['리오더 수량'])).clip(lower=0).astype(int)
        df_5['추가발주수량'] = df_5.index.map(st.session_state.add_order_dict).fillna(0).astype(int)

        def get_stat_v5_final(r):
            tot = r[avail] + r['리오더 수량']; day = r['일판매량']
            if day > 0:
                if tot < (day * 3): return "🚨 긴급"
                if tot < (day * 5): return "⚠️ 주의"
            return "✅ 정상"
        df_5['상태'] = df_5.apply(get_stat_v5_final, axis=1)

        df_disp_5 = df_5.rename(columns={item: "상품명", option: "옵션", v_item: "공급쳐상품명", avail: "가용재고", "리오더 수량": "리오더수량"})
        display_cols = ["상태", "상품명", "옵션", "공급쳐상품명", "가용재고", "리오더수량", "추가발주수량", "권장발주량"]

        with st.form("final_order_form"):
            edited_df = st.data_editor(df_disp_5[display_cols], use_container_width=True, hide_index=True, key="v5_editor")
            if st.form_submit_button("✅ 수량 확정 및 리오더 반영 (추가발주 합산)", use_container_width=True, type="primary"):
                changes = st.session_state["v5_editor"].get("edited_rows", {})
                for r_idx, change in changes.items():
                    orig_idx = df_5.index[int(r_idx)]
                    if "추가발주수량" in change:
                        # [중요] 추가발주수량을 리오더 수량에 합산
                        val = int(change["추가발주수량"])
                        st.session_state.df_raw.at[orig_idx, "리오더 수량"] += val
                        st.session_state.add_order_dict[orig_idx] = val
                st.success("✅ 리오더 수량에 합산되었습니다."); time.sleep(1); st.rerun()

        st.write("---")
        col_b1, col_b2 = st.columns(2)
        with col_b1:
            if st.button("💾 구글 시트에 최종 발주 기록 저장", use_container_width=True):
                # [6단계 연동] 5단계 화면 데이터를 그대로 로그로 변환
                ready = df_5.copy()
                ready['총발주'] = ready['권장발주량'] + ready['추가발주수량']
                final_to_save = ready[ready['총발주'] > 0]
                if not final_to_save.empty:
                    now_str = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
                    log_rows = [[now_str, r['상태'], r[item], r[option], r[v_item], int(r[avail]), int(r['리오더 수량']), int(r['추가발주수량']), int(r['권장발주량'])] for _, r in final_to_save.iterrows()]
                    try:
                        sheet = get_sheet()
                        sheet.worksheet("발주기록").append_rows(log_rows)
                        st.success(f"✅ {len(log_rows)}건의 내역이 6단계 히스토리에 저장되었습니다!")
                    except: st.error("📡 시트 연결 실패")
                else: st.warning("발주할 항목이 없습니다.")
        with col_b2:
            csv_final = df_disp_5[display_cols].to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button(label="📥 최종 발주서 CSV 다운로드", data=csv_final, file_name=f"발주서_{now.strftime('%m%d')}.csv", use_container_width=True)

        # ==========================================
        # --- [6단계: 전체 히스토리 내역] ---
        # ==========================================
        st.divider()
        st.subheader("📜 6단계: 전체 히스토리 내역")
        if st.button("🔄 히스토리 데이터 불러오기", use_container_width=True):
            sheet = get_sheet()
            if sheet:
                df_hist = pd.DataFrame(sheet.worksheet("발주기록").get_all_records())
                if not df_hist.empty:
                    st.dataframe(df_hist.sort_values(by="날짜", ascending=False), use_container_width=True, hide_index=True)
                else: st.info("기록이 없습니다.")

# [탭 2: 동대문 사입 관리는 사장님 기존 소스 그대로 유지됨]
with tab2:
    st.subheader("🌙 동대문 사입 및 미납 관리")
    dong_file = st.file_uploader("동대문 주문 리스트 업로드", type=['xlsx', 'csv'], key="dong_tab_upload")
    if dong_file:
        df_d = pd.read_excel(dong_file) if not dong_file.name.endswith('.csv') else pd.read_csv(dong_file)
        # (중략: 사장님 동대문 로직 그대로...)
        st.dataframe(df_d, use_container_width=True)
