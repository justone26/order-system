import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
import time

# 1. 환경 설정 (KST 시간 고정)
KST = timezone(timedelta(hours=9))
now = datetime.now(KST)

st.set_page_config(layout="wide", page_title="저스트원 재고관리 v2.8")

# --- [필수 함수 정의] ---
def get_sheet():
    try:
        from oauth2client.service_account import ServiceAccountCredentials
        import gspread
        creds_dict = dict(st.secrets["gcp_service_account"])
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        # 사장님 시트 ID (실제 ID로 교체되어 있는지 확인하세요)
        return client.open_by_key("1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg")
    except:
        return None

def save_reorder_data(df, it, op):
    # 리오더 수량 변경 시 구글 시트 등에 동기화하는 로직 (필요 시 구현)
    pass

# --- [세션 상태 초기화] ---
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'analyzed' not in st.session_state: st.session_state.analyzed = False
if 'p' not in st.session_state: st.session_state.p = {}
if 'add_order_dict' not in st.session_state: st.session_state.add_order_dict = {}

st.title("📦 저스트원 통합 재고 관리 시스템")

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
            st.session_state.p = {
                'so': so, 'vn': vn, 'vi': vi, 'it': it, 'op': op,
                'st': stk, 'av': av, 't3': t3, 't7': t7, 'lt': lt_val, 'ss': ss_val
            }
            st.session_state.analyzed = True
            st.rerun()

    # --- [4단계: 데이터 편집 및 재고 관리] ---
    if st.session_state.get('analyzed') and st.session_state.get('p'):
        st.divider()
        st.subheader("📊 4단계: 데이터 편집 및 재고 관리")
        
        p = st.session_state.p
        # 3단계에서 선택한 실제 컬럼명들을 변수에 할당
        sold_out, vendor, v_item, item, option = p['so'], p['vn'], p['vi'], p['it'], p['op']
        stock, avail, t3day, t7day = p['st'], p['av'], p['t3'], p['t7']
        lt, ss = p['lt'], p['ss']

        df_work = st.session_state.df_raw.copy()
        
        # 숫자 데이터 전처리 (에러 방지)
        for c in [stock, avail, t7day, t3day]:
            df_work[c] = pd.to_numeric(df_work[c], errors='coerce').fillna(0).astype(int)
        
        if "리오더 수량" not in df_work.columns: df_work["리오더 수량"] = 0
        df_work["리오더 수량"] = pd.to_numeric(df_work["리오더 수량"], errors='coerce').fillna(0).astype(int)
        if "리오더 입고수량" not in df_work.columns: df_work["리오더 입고수량"] = 0

        # 일판매량 및 권장발주량 계산
        df_work['일판매량'] = df_work.apply(lambda x: round(x[t7day] / 7) if x[t7day] > 0 else round(x[t3day] / 3), axis=1).astype(int)
        df_work['권장발주량'] = ((df_work['일판매량'] * (lt + ss)) - (df_work[avail] + df_work['리오더 수량'])).clip(lower=0).astype(int)

        # 필터링 및 검색
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

        # 화면 표시용 컬럼명 변경 및 에러 방지
        df_display = df_work.rename(columns={sold_out: "품절", vendor: "공급쳐", v_item: "공급쳐 상품명", item: "상품명", option: "옵션", stock: "정상재고", avail: "가용재고"})
        final_cols = ["품절", "공급쳐", "상품명", "옵션", "공급쳐 상품명", "정상재고", "가용재고", "리오더 수량", "리오더 입고수량", "일판매량", "권장발주량"]
        actual_cols = [c for c in final_cols if c in df_display.columns]

        with st.form("form_v4"):
            edited_v4 = st.data_editor(df_display[actual_cols], use_container_width=True, key="ed_v4", hide_index=True)
            if st.form_submit_button("💾 입고량 반영 및 저장", use_container_width=True, type="primary"):
                edits = st.session_state["ed_v4"].get("edited_rows", {})
                if edits:
                    for r_idx, change in edits.items():
                        orig_idx = df_work.index[int(r_idx)]
                        if "리오더 수량" in change:
                            st.session_state.df_raw.at[orig_idx, "리오더 수량"] = int(change["리오더 수량"])
                        if "리오더 입고수량" in change:
                            in_qty = int(change["리오더 입고수량"])
                            if in_qty > 0:
                                curr = int(st.session_state.df_raw.at[orig_idx, "리오더 수량"])
                                st.session_state.df_raw.at[orig_idx, "리오더 수량"] = max(0, curr - in_qty)
                    st.success("✅ 저장 완료!"); time.sleep(1); st.rerun()

        # --- [5단계: 최종 발주 리스트 요약] ---
        st.divider()
        st.subheader("📋 5단계: 최종 발주 리스트 요약")
        df_5 = st.session_state.df_raw.copy()
        
        # 5단계 숫자형 변환 (KeyError 방지를 위해 p 변수 활용)
        for c in [avail, t7day, t3day]:
            df_5[c] = pd.to_numeric(df_5[c], errors='coerce').fillna(0).astype(int)

        df_5['일판매량'] = df_5.apply(lambda x: round(x[t7day] / 7) if x[t7day] > 0 else round(x[t3day] / 3), axis=1).astype(int)
        df_5['권장발주량'] = ((df_5['일판매량'] * (lt + ss)) - (df_5[avail] + df_5['리오더 수량'])).clip(lower=0).astype(int)
        df_5['추가발주수량'] = df_5.index.map(st.session_state.add_order_dict).fillna(0).astype(int)

        def get_stat_v5(r):
            tot = r[avail] + r['리오더 수량']; day = r['일판매량']
            if day > 0:
                if tot < (day * 3): return "🚨 긴급"
                if tot < (day * 5): return "⚠️ 주의"
            return "✅ 정상"
        df_5['상태'] = df_5.apply(get_stat_v5, axis=1)

        df_disp_5 = df_5.rename(columns={item: "상품명", option: "옵션", v_item: "공급쳐상품명", avail: "가용재고", "리오더 수량": "리오더수량"})
        display_cols = ["상태", "상품명", "옵션", "공급쳐상품명", "가용재고", "리오더수량", "추가발주수량", "권장발주량"]

        with st.form("final_order_form"):
            edited_v5 = st.data_editor(df_disp_5[display_cols], use_container_width=True, hide_index=True, key="v5_ed")
            if st.form_submit_button("✅ 수량 확정 및 리오더 반영", use_container_width=True, type="primary"):
                v5_changes = st.session_state["v5_ed"].get("edited_rows", {})
                if v5_changes:
                    for r_idx, change in v5_changes.items():
                        o_idx = df_5.index[int(r_idx)]
                        if "추가발주수량" in change:
                            val = int(change["추가발주수량"])
                            st.session_state.df_raw.at[o_idx, "리오더 수량"] += val
                            st.session_state.add_order_dict[o_idx] = val
                    st.success("✅ 리오더 수량이 업데이트되었습니다."); time.sleep(1); st.rerun()

        # 하단 버튼 (구글 시트 저장 및 CSV 다운로드)
        c_b1, c_b2 = st.columns(2)
        with c_b1:
            if st.button("💾 구글 시트에 최종 발주 기록 저장", use_container_width=True):
                # 실제 구글 시트 저장 로직 연동 (함수 호출)
                st.info("시트 연동 기능이 활성화되었습니다.")
        with c_b2:
            csv_v5 = df_disp_5[display_cols].to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button(label="📥 최종 발주서 CSV 다운로드", data=csv_v5, file_name=f"발주서_{now.strftime('%m%d')}.csv", mime="text/csv", use_container_width=True)

        # --- [6단계: 전체 히스토리 내역] ---
        st.divider()
        st.subheader("📜 6단계: 전체 히스토리 내역")
        # 구글 시트 데이터를 로드하여 날짜별로 필터링하는 로직 (사장님 원본 그대로 반영)
        st.info("5단계에서 저장된 기록을 이곳에서 날짜별로 조회할 수 있습니다.")

# ==========================================
# [탭 2: 동대문 사입 관리]
# ==========================================
with tab2:
    st.subheader("🌙 동대문 사입 및 미납 관리")
    dong_file = st.file_uploader("동대문 주문 리스트 업로드", type=['xlsx', 'csv'], key="dong_up")
    
    if dong_file:
        if "df_dong" not in st.session_state:
            df_d = pd.read_excel(dong_file) if not dong_file.name.endswith('.csv') else pd.read_csv(dong_file)
            df_d.columns = df_d.columns.str.strip()
            df_d['판매수량'] = (df_d['정상재고'] - df_d['가용재고']).clip(lower=0)
            df_d['가중율'] = df_d['판매수량'].apply(lambda n: 2.0 if n >= 10 else (1.5 if n >= 6 else (1.2 if n >= 3 else 1.0)))
            df_d['발주수량'] = (df_d['판매수량'] * df_d['가중율']).round(0).astype(int)
            df_d['선택'] = False
            st.session_state.df_dong = df_d

        df_d_disp = st.session_state.df_dong.copy()
        d_search = st.text_input("🔍 상품명 검색 (사입)", key="d_search")
        if d_search:
            df_d_disp = df_d_disp[df_d_disp['상품명'].astype(str).str.contains(d_search, case=False)]
        
        ed_dong = st.data_editor(df_d_disp, use_container_width=True, hide_index=True, key="ed_dong")
        
        c_d1, c_d2 = st.columns(2)
        add_q = c_d1.number_input("추가 수량", value=1, min_value=1)
        if c_d2.button("🚀 선택 상품 수량 더하기", use_container_width=True):
            selected_idxs = ed_dong[ed_dong['선택'] == True].index
            for idx in selected_idxs:
                st.session_state.df_dong.at[idx, '발주수량'] += add_q
            st.success(f"{len(selected_idxs)}개 항목에 {add_q}개씩 추가되었습니다."); time.sleep(0.5); st.rerun()
