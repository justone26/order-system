import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
import time
import io

# 1. [환경 설정]
KST = timezone(timedelta(hours=9))
now = datetime.now(KST)
st.set_page_config(layout="wide", page_title="저스트원 재고관리 v4.0")

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

# --- [세션 상태 초기화] ---
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'analyzed' not in st.session_state: st.session_state.analyzed = False
if 'p' not in st.session_state: st.session_state.p = {}
if 'add_order_dict' not in st.session_state: st.session_state.add_order_dict = {}

st.title("📦 저스트원 통합 재고 관리 v4.0")

tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 관리"])

with tab1:
    # --- 1~3단계: 설정 ---
    st.subheader("📁 1~3단계: 데이터 업로드 및 분석 설정")
    up_file = st.file_uploader("엑셀/CSV 파일 업로드", type=['xlsx', 'xls', 'csv'], key="main_up")
    
    if st.button("🔄 시스템 전체 초기화", use_container_width=True):
        for key in list(st.session_state.keys()): del st.session_state[key]
        st.rerun()

    if up_file and st.session_state.df_raw is None:
        df = pd.read_csv(up_file) if up_file.name.endswith('.csv') else pd.read_excel(up_file)
        df.columns = df.columns.str.strip()
        if "리오더 수량" not in df.columns: df["리오더 수량"] = 0
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
            so = st.selectbox("품절 여부", cols, index=auto_idx(['품절'])); vn = st.selectbox("공급처", cols, index=auto_idx(['공급처']))
            vi = st.selectbox("공급처 상품명", cols, index=auto_idx(['공급처상품명']))
        with c2:
            it = st.selectbox("상품명", cols, index=auto_idx(['상품명'])); op = st.selectbox("옵션", cols, index=auto_idx(['옵션']))
            stk = st.selectbox("정상재고", cols, index=auto_idx(['정상재고']))
        with c3:
            av = st.selectbox("가용재고", cols, index=auto_idx(['가용재고'])); t3 = st.selectbox("3일 판매", cols, index=auto_idx(['3일']))
            t7 = st.selectbox("7일 판매", cols, index=auto_idx(['7일']))
        
        lt_val = st.number_input("⏳ 리드타임 (일)", value=7)
        ss_val = st.number_input("🛡️ 안전재고 (일)", value=3)

        if st.button("🚀 데이터 분석 시작", use_container_width=True, type="primary"):
            st.session_state.p = {'so': so, 'vn': vn, 'vi': vi, 'it': it, 'op': op, 'st': stk, 'av': av, 't3': t3, 't7': t7, 'lt': lt_val, 'ss': ss_val}
            st.session_state.analyzed = True
            st.rerun()

    # --- [4단계: 리오더 차감 관리] ---
    if st.session_state.get('analyzed'):
        st.divider(); st.subheader("📊 4단계: 데이터 편집 및 재고 관리")
        p = st.session_state.p
        df_work = st.session_state.df_raw.copy()
        
        if "리오더 수량" not in df_work.columns: df_work["리오더 수량"] = 0
        for col_name in [p['st'], p['av'], p['t3'], "리오더 수량"]:
            df_work[col_name] = pd.to_numeric(df_work[col_name], errors='coerce').fillna(0).astype(int)
        
        df_work["리오더 입고수량"] = 0
        df_work['일판매량'] = (df_work[p['t3']] / 3).round(1)
        df_work['권장발주량'] = ((df_work['일판매량'] * (p['lt'] + p['ss'])) - (df_work[p['av']] + df_work['리오더 수량'])).clip(lower=0).astype(int)

        f4_c1, f4_c2, f4_c3 = st.columns([2, 1, 1])
        s4_q = f4_c1.text_input("🔍 4단계 검색", key="s4_search")
        m4_f = f4_c2.selectbox("4단계 필터", ["전체보기", "정상만", "품절만"], index=1, key="m4_filter")
        d4_h = f4_c3.date_input("🗓️ 기록 날짜", now.date(), key="d4_date")

        if m4_f == "정상만": df_work = df_work[~df_work[p['so']].astype(str).str.contains('품절', na=False)]
        elif m4_f == "품절만": df_work = df_work[df_work[p['so']].astype(str).str.contains('품절', na=False)]
        if s4_q: df_work = df_work[df_work[p['it']].astype(str).str.contains(s4_q, case=False) | df_work[p['op']].astype(str).str.contains(s4_q, case=False)]

        df_disp4 = df_work.rename(columns={p['so']:"품절", p['vn']:"공급쳐", p['it']:"상품명", p['op']:"옵션", p['av']:"가용재고", p['t3']:"3일 판매 합계"})
        cols4 = ["품절", "공급쳐", "상품명", "옵션", "가용재고", "리오더 수량", "리오더 입고수량", "3일 판매 합계", "일판매량", "권장발주량"]

        with st.form("form_v4"):
            ed4 = st.data_editor(df_disp4[cols4], use_container_width=True, hide_index=True, key="ed4")
            if st.form_submit_button("💾 4단계 입고 반영 (차감)"):
                edits = st.session_state["ed4"].get("edited_rows", {})
                for r_idx, change in edits.items():
                    o_idx = df_work.index[int(r_idx)]
                    if "리오더 입고수량" in change:
                        in_qty = int(change["리오더 입고수량"])
                        st.session_state.df_raw.at[o_idx, "리오더 수량"] = max(0, int(st.session_state.df_raw.at[o_idx, "리오더 수량"]) - in_qty)
                st.success("✅ 차감 완료!"); time.sleep(0.5); st.rerun()

        # --- [5단계: 최종 발주 및 엑셀 다운로드] ---
# 1. 안전 장치: 분석이 완료되었고 데이터가 있을 때만 실행
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📋 5단계: 최종 발주 리스트 요약")

    # [변수 선언] 세션에 저장된 매칭 정보 사용
    p = st.session_state.p
    avail = p['av']
    t7day = p['t7']
    t3day = p['t3']
    item = p['it']
    option = p['op']
    v_item = p['vi']
    lt = p['lt']
    ss = p['ss']

    # 데이터 복사 및 전처리
    df_5 = st.session_state.df_raw.copy()
    
    # 숫자 데이터 변환 (계산 에러 방지)
    for c in [avail, '리오더 수량', t7day, t3day]:
        if c in df_5.columns:
            df_5[c] = pd.to_numeric(df_5[c], errors='coerce').fillna(0).astype(int)

    # --- [판매량 및 발주량 핵심 로직 적용] ---
    # 1. 일판매량: 7일 데이터가 있으면 7일 우선, 없으면 3일 기준
    df_5['일판매량'] = df_5.apply(lambda x: round(x[t7day] / 7) if x[t7day] > 0 else round(x[t3day] / 3), axis=1).astype(int)
    # 2. 권장발주량: (일판매량 * 확보일수) - (현재고 + 리오더중 수량)
    df_5['권장발주량'] = ((df_5['일판매량'] * (lt + ss)) - (df_5[avail] + df_5['리오더 수량'])).clip(lower=0).astype(int)
    # 3. 추가발주수량: 사장님이 수동으로 입력한 값 매칭
    df_5['추가발주수량'] = df_5.index.map(st.session_state.add_order_dict).fillna(0).astype(int)
    # 4. 최종발주합계: 시스템 권장량 + 사장님 추가량
    df_5['최종발주합계'] = df_5['권장발주량'] + df_5['추가발주수량']

    # 상태 판별 (긴급/주의/정상)
    def get_stat_v5_final(r):
        tot = r[avail] + r['리오더 수량']
        day = r['일판매량']
        if day > 0:
            if tot < (day * 3): return "🚨 긴급"
            if tot < (day * 5): return "⚠️ 주의"
        return "✅ 정상"
    df_5['상태'] = df_5.apply(get_stat_v5_final, axis=1)

    # 화면 표시용 컬럼 이름 변경 및 순서 배치
    df_disp_5 = df_5.rename(columns={item: "상품명", option: "옵션", v_item: "공급쳐상품명", avail: "가용재고", "리오더 수량": "리오더수량"})
    
    # 사장님이 말씀하신 "화면에 나와야 되는 애들" 위주로 배치
    display_cols = ["상태", "상품명", "옵션", "가용재고", "리오더수량", "일판매량", "권장발주량", "추가발주수량", "최종발주합계"]

    # 5단계 검색 및 필터 (필요시 사용)
    f5_c1, f5_c2 = st.columns([2, 1])
    s5_q = f5_c1.text_input("🔍 상품명/옵션 검색", key="s5_search_v43")
    m5_f = f5_c2.selectbox("발주 상태 필터", ["전체보기", "정상만", "품절만"], index=1, key="m5_filter_v43")

    # 필터 적용 로직
    if m5_f == "정상만": df_disp_5 = df_disp_5[~df_5[p['so']].astype(str).str.contains('품절', na=False)]
    elif m5_f == "품절만": df_disp_5 = df_disp_5[df_5[p['so']].astype(str).str.contains('품절', na=False)]
    if s5_q: df_disp_5 = df_disp_5[df_disp_5["상품명"].astype(str).str.contains(s5_q, case=False) | df_disp_5["옵션"].astype(str).str.contains(s5_q, case=False)]

    # 2. 데이터 에디터 (추가발주수량 수정 가능)
    with st.form("final_order_form_v43"):
        edited_df = st.data_editor(
            df_disp_5[display_cols],
            use_container_width=True,
            hide_index=True,
            key="v5_editor_v43",
            column_config={
                "최종발주합계": st.column_config.NumberColumn("최종발주합계", format="%d", disabled=True),
                "추가발주수량": st.column_config.NumberColumn("추가발주수량", format="%d", help="수동 발주량을 입력하세요.")
            }
        )
        
        if st.form_submit_button("✅ 수량 확정 및 리오더 반영", use_container_width=True, type="primary"):
            changes = st.session_state["v5_editor_v43"].get("edited_rows", {})
            if changes:
                for r_idx, change in changes.items():
                    orig_idx = df_disp_5.index[int(r_idx)]
                    if "추가발주수량" in change:
                        val = int(change["추가발주수량"])
                        # 원본 데이터의 리오더 수량에 합산
                        st.session_state.df_raw.at[orig_idx, "리오더 수량"] += val
                        st.session_state.add_order_dict[orig_idx] = val
                st.success("✅ 발주 수량이 리오더에 합산되었습니다.")
                time.sleep(1); st.rerun()

    # 3. 하단 버튼 구역 (저장 및 CSV 다운로드)
    st.write("---")
    col_b1, col_b2 = st.columns(2)

    with col_b1:
        if st.button("💾 구글 시트에 최종 발주 기록 저장", use_container_width=True):
            # 최종 합계가 0보다 큰 것만 필터링해서 저장
            ready_to_save = df_disp_5[df_disp_5['최종발주합계'] > 0]
            if not ready_to_save.empty:
                now_str = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
                log_rows = [[now_str, r['상태'], r['상품명'], r['옵션'], r['가용재고'], r['리오더수량'], r['추가발주수량'], r['권장발주량'], r['최종발주합계']] for _, r in ready_to_save.iterrows()]
                try:
                    sheet = get_sheet()
                    sheet.worksheet("발주기록").append_rows(log_rows)
                    st.success(f"✅ {len(log_rows)}건 저장 완료!")
                    st.session_state.add_order_dict = {} # 저장 후 딕셔너리 초기화
                    time.sleep(1); st.rerun()
                except Exception as e:
                    st.error(f"📡 저장 실패: {e}")

    with col_b2:
        # 다운로드용 데이터도 합계가 있는 것만
        csv_data = df_disp_5[df_disp_5['최종발주합계'] > 0]
        if not csv_data.empty:
            csv_file = csv_data[display_cols].to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button(
                label="📥 최종 발주서 CSV 다운로드",
                data=csv_file,
                file_name=f"저스트원_발주서_{datetime.now(KST).strftime('%m%d')}.csv",
                mime="text/csv",
                use_container_width=True
            )
        # --- [6단계: 히스토리] ---
        st.divider(); st.subheader("📜 6단계: 전체 히스토리 내역")
        if st.button("🔄 히스토리 새로고침", use_container_width=True):
            sheet = get_sheet()
            if sheet:
                df_hist = pd.DataFrame(sheet.worksheet("발주기록").get_all_records())
                st.dataframe(df_hist.sort_values(by=df_hist.columns[0], ascending=False), use_container_width=True, hide_index=True)
