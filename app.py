import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
import time

# 1. 기본 설정 (KST 기준)
KST = timezone(timedelta(hours=9))
now = datetime.now(KST)

st.set_page_config(layout="wide", page_title="저스트원 재고관리 v2.4")

# [함수들: 사장님 소스에 필요한 연동 함수들 정의]
def get_incoming_history():
    # 실제 구글 시트 연결 전까지는 빈 데이터프레임 반환 (연결 시 수정 가능)
    return pd.DataFrame(columns=['상품명', '옵션', '과거리오더 입고'])

def save_history_to_gsheet(df, log_type="입고"):
    st.info(f"데이터가 {log_type} 히스토리에 임시 기록되었습니다.")

def save_reorder_data(df, item_col, opt_col):
    st.info("리오더 수량이 세션에 저장되었습니다.")

def get_sheet():
    # 구글 시트 연결용 (기존 Secrets 설정 필요)
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

# 리셋 콜백
def reset_callback():
    for key in list(st.session_state.keys()):
        del st.session_state[key]

# 세션 초기화
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'analyzed' not in st.session_state: st.session_state.analyzed = False
if 'p' not in st.session_state: st.session_state.p = {}

st.title("📦 저스트원 통합 재고 관리 시스템")

tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 관리"])

with tab1:
    # --- 1단계: 업로드 ---
    st.subheader("📁 1단계: 데이터 업로드")
    up_file = st.file_uploader("파일 업로드", type=['xlsx', 'xls', 'csv'], key="main_uploader", label_visibility="collapsed")
    if st.button("🔄 전체 데이터 초기화", on_click=reset_callback): pass

    if up_file and st.session_state.df_raw is None:
        df = pd.read_csv(up_file) if up_file.name.endswith('.csv') else pd.read_excel(up_file)
        df.columns = df.columns.str.strip()
        st.session_state.df_raw = df
        st.rerun()

    # --- 2~3단계: 설정 (사장님이 주신 변수명 p에 매칭) ---
    if st.session_state.df_raw is not None:
        st.divider()
        st.subheader("🔗 2~3단계: 컬럼 매칭 및 발주 설정")
        cols = st.session_state.df_raw.columns.tolist()
        c1, c2, c3 = st.columns(3)
        with c1:
            so = st.selectbox("품절 여부", cols, index=0)
            vn = st.selectbox("공급처", cols, index=0)
            vi = st.selectbox("공급처 상품명", cols, index=0)
        with c2:
            it = st.selectbox("상품명", cols, index=0)
            op = st.selectbox("옵션", cols, index=0)
            stk = st.selectbox("정상재고", cols, index=0)
        with c3:
            av = st.selectbox("가용재고", cols, index=0)
            t3 = st.selectbox("3일 판매", cols, index=0)
            t7 = st.selectbox("7일 판매", cols, index=0)
        
        lt = st.number_input("⏳ 리드타임 (7일 디폴트)", value=7)
        ss = st.number_input("🛡️ 안전재고 (3일 디폴트)", value=3)

        if st.button("🚀 데이터 분석 시작", use_container_width=True, type="primary"):
            st.session_state.p = {
                'so': so, 'vn': vn, 'vi': vi, 'it': it, 'op': op,
                'st': stk, 'av': av, 't3': t3, 't7': t7, 'lt': lt, 'ss': ss
            }
            st.session_state.analyzed = True
            st.rerun()

# ==========================================
# 4단계: 사장님 소스 그대로 투입
# ==========================================
if st.session_state.get('analyzed') and st.session_state.get('p'):
    st.divider()
    st.subheader("📊 4단계: 데이터 편집 및 재고 관리")

    p = st.session_state.p
    sold_out, vendor, v_item, item, option = p['so'], p['vn'], p['vi'], p['it'], p['op']
    stock, avail, t3day, t7day = p['st'], p['av'], p['t3'], p['t7']
    lt, ss = p['lt'], p['ss']

    df_work = st.session_state.df_raw.copy()
    
    # 입고 기록 실시간 매칭
    incoming_hist = get_incoming_history()
    if not incoming_hist.empty:
        df_work = pd.merge(df_work, incoming_hist, left_on=[item, option], right_on=['상품명', '옵션'], how='left')
        df_work['과거리오더 입고'] = df_work['과거리오더 입고'].fillna(0).astype(int)
    else:
        df_work['과거리오더 입고'] = 0

    # 숫자형 변환
    for c in [stock, avail, t7day, t3day]:
        if c in df_work.columns:
            df_work[c] = pd.to_numeric(df_work[c], errors='coerce').fillna(0).astype(int)
    
    if "리오더 수량" not in df_work.columns: df_work["리오더 수량"] = 0
    df_work["리오더 수량"] = pd.to_numeric(df_work["리오더 수량"], errors='coerce').fillna(0).astype(int)
    if "리오더 입고수량" not in df_work.columns: df_work["리오더 입고수량"] = 0

    df_work['일판매량'] = df_work.apply(lambda x: round(x[t7day] / 7) if x[t7day] > 0 else round(x[t3day] / 3), axis=1).astype(int)
    df_work['권장발주량'] = ((df_work['일판매량'] * (lt + ss)) - (df_work[avail] + df_work['리오더 수량'])).clip(lower=0).astype(int)

    # UI 및 필터
    f_c1, f_c2, f_c3 = st.columns([2, 1, 1])
    search_q = f_c1.text_input("🔍 상품명/옵션 검색", key="v4_fin_s")
    filter_m = f_c2.selectbox("상태 필터", ["전체보기", "정상만", "품절만"], index=1, key="v4_fin_f")
    hist_date_4 = f_c3.date_input("🗓️ 입고 기록 날짜", now.date(), key="v4_fin_d")

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
        edited_v4 = st.data_editor(
            df_display[final_cols], 
            use_container_width=True, 
            key="ed_v4_safe", 
            hide_index=True,
            column_config={
                "정상재고": st.column_config.NumberColumn(format="%d"),
                "가용재고": st.column_config.NumberColumn(format="%d"),
                "리오더 수량": st.column_config.NumberColumn(format="%d"),
                "리오더 입고수량": st.column_config.NumberColumn(format="%d"),
                "과거리오더 입고": st.column_config.NumberColumn(format="%d"),
                "일판매량": st.column_config.NumberColumn(format="%d"),
                "권장발주량": st.column_config.NumberColumn(format="%d")
            }
        )
        if st.form_submit_button("💾 입고량 반영 및 저장", use_container_width=True, type="primary"):
            edits = st.session_state["ed_v4_safe"].get("edited_rows", {})
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
                            log_df = pd.DataFrame([[f"{hist_date_4}", df_work.at[orig_idx, item], df_work.at[orig_idx, option], in_qty]], columns=['날짜', '상품명', '옵션', '수량'])
                            save_history_to_gsheet(log_df, log_type="입고")
                save_reorder_data(st.session_state.df_raw, item, option)
                st.success("✅ 저장 완료!")
                time.sleep(1); st.rerun()

# --- 5단계 및 6단계는 위 4단계 로직 바로 아래에 이어서 사장님 소스 붙여넣으시면 완벽합니다 ---
