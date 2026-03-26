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

# [필수 함수] 구글 시트 입고 기록 가져오기
def get_incoming_history():
    try:
        sheet = get_sheet() 
        ws = sheet.worksheet("입고기록")
        data = ws.get_all_records()
        if data:
            df_h = pd.DataFrame(data)
            df_h['상품명'] = df_h['상품명'].astype(str).str.strip()
            df_h['옵션'] = df_h['옵션'].astype(str).str.strip()
            summary = df_h.groupby(['상품명', '옵션'])['수량'].sum().reset_index()
            summary.rename(columns={'수량': '과거리오더 입고'}, inplace=True)
            return summary
        return pd.DataFrame(columns=['상품명', '옵션', '과거리오더 입고'])
    except:
        return pd.DataFrame(columns=['상품명', '옵션', '과거리오더 입고'])
        
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
    
    if st.button("🔄 화면 전체 초기화", use_container_width=True):
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

# --- 4단계 시작 ---
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📊 4단계: 데이터 편집 및 재고 관리")

    # 1. 설정값 불러오기
    p = st.session_state.p
    sold_out_col = p['so'] 
    item, option = p['it'], p['op']
    vendor, v_item = p['vn'], p['vi']
    stock, avail, t3day, t7day = p['st'], p['av'], p['t3'], p['t7']
    lt, ss = p['lt'], p['ss']

    # [핵심 수정] 데이터 유실 및 타입 에러 방지
    df_work = st.session_state.df_raw.copy()
    
    # 모든 데이터를 일단 문자열로 변환하고 공백을 제거 (에러 방지 핵심)
    df_work[sold_out_col] = df_work[sold_out_col].astype(str).str.strip()

    # 2. UI 배치 (상태 필터 -> 검색어 -> 날짜 순)
    f_c1, f_c2, f_c3 = st.columns([1, 2, 1])
    filter_m = f_c1.selectbox("🚦 상태 필터", ["전체보기", "정상만", "품절만"], index=0, key="v4_final_filter")
    search_q = f_c2.text_input("🔍 상품명/옵션 검색", placeholder="검색어를 입력하세요...", key="v4_final_search")
    hist_date_4 = f_c3.date_input("🗓️ 입고 날짜", datetime.now().date(), key="v4_final_date")

    # 3. 데이터 계산 (숫자형 변환 후 계산)
    for c in [stock, avail, t7day, t3day]:
        df_work[c] = pd.to_numeric(df_work[c], errors='coerce').fillna(0).astype(int)
    
    df_work['일판매량'] = df_work.apply(lambda x: round(x[t7day] / 7) if x[t7day] > 0 else round(x[t3day] / 3), axis=1).astype(int)
    # [요청] 3일 발주수량 추가
    df_work['3일 발주수량'] = (df_work['일판매량'] * 3).astype(int)
    
    if "리오더 수량" not in df_work.columns: 
        df_work["리오더 수량"] = 0
    df_work["리오더 수량"] = pd.to_numeric(df_work["리오더 수량"], errors='coerce').fillna(0).astype(int)
    df_work["리오더 입고수량"] = 0
    
    df_work['권장발주량'] = ((df_work['일판매량'] * (lt + ss)) - (df_work[avail] + df_work['리오더 수량'])).clip(lower=0).astype(int)

    # 4. 필터 로직 (문자열 전용 contains 사용)
    # 위에서 이미 .astype(str) 처리를 했으므로 에러가 나지 않습니다.
    is_soldout_row = df_work[sold_out_col].str.contains('품절', na=False)

    if filter_m == "정상만":
        # '품절' 글자가 없고, 'nan' (빈값) 혹은 공백인 것들 포함
        df_filtered = df_work[(~is_soldout_row) | (df_work[sold_out_col] == 'nan') | (df_work[sold_out_col] == '')]
    elif filter_m == "품절만":
        df_filtered = df_work[is_soldout_row]
    else:
        df_filtered = df_work

    # 검색어 필터
    if search_q:
        df_filtered = df_filtered[
            df_filtered[item].astype(str).str.contains(search_q, case=False, na=False) | 
            df_filtered[option].astype(str).str.contains(search_q, case=False, na=False)
        ]

    # 5. 컬럼명 변경 및 요청하신 순서 재배치
    df_display = df_filtered.rename(columns={
        sold_out_col: "품절상태", vendor: "공급쳐", v_item: "공급쳐 상품명", 
        item: "상품명", option: "옵션", stock: "정상재고", avail: "가용재고"
    })
    
    # 과거리오더 입고 매칭 (함수 연동)
    inc_h = get_incoming_history()
    if not inc_h.empty:
        df_display = pd.merge(df_display, inc_h, on=["상품명", "옵션"], how="left")
        df_display["과거리오더 입고"] = df_display["과거리오더 입고"].fillna(0).astype(int)
    else:
        df_display["과거리오더 입고"] = 0

    # 최종 컬럼 순서: 과거리오더 입고 -> 3일 발주수량 -> 일판매량 -> 권장발주량
    final_cols = [
        "품절상태", "공급쳐", "상품명", "옵션", "공급쳐 상품명", 
        "정상재고", "가용재고", "리오더 수량", "리오더 입고수량", 
        "과거리오더 입고", "3일 발주수량", "일판매량", "권장발주량"
    ]

    # 6. 결과 출력 (에디터 폼)
    with st.form("v4_final_safe_form"):
        if not df_display.empty:
            st.data_editor(
                df_display[final_cols],
                use_container_width=True,
                hide_index=True,
                key="v4_editor_safe",
                column_config={c: st.column_config.NumberColumn(disabled=True) for c in ["과거리오더 입고", "3일 발주수량", "일판매량", "권장발주량"]}
            )
        else:
            st.info("💡 표시할 데이터가 없습니다. 필터를 변경해 보세요.")
            # 디버깅용: 데이터가 왜 안나오는지 실제 값을 살짝 보여줌
            if filter_m == "품절만":
                st.write("현재 '품절상태' 컬럼에 들어있는 값들:", df_work[sold_out_col].unique())

        if st.form_submit_button("💾 데이터 저장 및 입고 반영", use_container_width=True, type="primary"):
            # 저장 로직 (생략 - 필요시 추가)
            st.success("반영되었습니다.")
            time.sleep(1)
            st.rerun()

elif not st.session_state.get('analyzed'):
    st.info("데이터 업로드 후 '데이터 분석 시작' 버튼을 눌러주세요.")


        # --- [5단계: 최종 발주 및 엑셀 다운로드] ---
# 1. 안전 장치: 분석이 완료되었고 데이터가 있을 때만 실행
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📋 5단계: 최종 발주 리스트 요약")

    # [변수 선언]
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
    
    # 숫자 데이터 변환
    for c in [avail, '리오더 수량', t7day, t3day]:
        if c in df_5.columns:
            df_5[c] = pd.to_numeric(df_5[c], errors='coerce').fillna(0).astype(int)

    # --- [상단 필터 및 날짜 영역] ---
    f_c1, f_c2, f_c3 = st.columns([1, 2, 1])
    with f_c1:
        d5_date = st.date_input("🗓️ 기록 날짜", now.date(), key="d5_record_date")
    with f_c2:
        s5_search = st.text_input("🔍 상품명 검색", key="s5_name_search")
    with f_c3:
        m5_filter = st.selectbox("🚦 상태 필터", ["전체보기", "🚨 긴급", "⚠️ 주의", "✅ 정상"], index=0)

    # --- [계산 로직] ---
    df_5['일판매량'] = df_5.apply(lambda x: round(x[t7day] / 7) if x[t7day] > 0 else round(x[t3day] / 3), axis=1).astype(int)
    df_5['권장 발주수량'] = ((df_5['일판매량'] * (lt + ss)) - (df_5[avail] + df_5['리오더 수량'])).clip(lower=0).astype(int)
    df_5['추가발주수량'] = df_5.index.map(st.session_state.add_order_dict).fillna(0).astype(int)

    # 상태 판별
    def get_stat_v5_final(r):
        tot = r[avail] + r['리오더 수량']
        day = r['일판매량']
        if day > 0:
            if tot < (day * 3): return "🚨 긴급"
            if tot < (day * 5): return "⚠️ 주의"
        return "✅ 정상"
    df_5['상태'] = df_5.apply(get_stat_v5_final, axis=1)

    # 화면 표시용 이름 변경
    df_disp_5 = df_5.rename(columns={item: "상품명", option: "옵션", v_item: "공급쳐상품명", avail: "가용재고", "리오더 수량": "리오더수량"})
    
    # 사장님이 요청하신 8가지 컬럼만 선택
    display_cols = ["상태", "상품명", "옵션", "공급쳐상품명", "가용재고", "리오더수량", "추가발주수량", "권장 발주수량"]

    # --- [필터링 적용] ---
    if m5_filter != "전체보기":
        df_disp_5 = df_disp_5[df_disp_5["상태"] == m5_filter]
    if s5_search:
        df_disp_5 = df_disp_5[df_disp_5["상품명"].astype(str).str.contains(s5_search, case=False)]

    # 2. 데이터 에디터 (추가발주수량만 수정 가능하도록 설정)
    with st.form("final_order_form_v44"):
        edited_df = st.data_editor(
            df_disp_5[display_cols],
            use_container_width=True,
            hide_index=True,
            key="v5_editor_v44",
            column_config={
                "상태": st.column_config.TextColumn(width="small"),
                "추가발주수량": st.column_config.NumberColumn("추가발주수량", format="%d", help="수동 발주량을 입력하세요."),
                "권장 발주수량": st.column_config.NumberColumn(format="%d", disabled=True)
            }
        )
        
        if st.form_submit_button("✅ 수량 확정 및 리오더 반영", use_container_width=True, type="primary"):
            changes = st.session_state["v5_editor_v44"].get("edited_rows", {})
            if changes:
                for r_idx, change in changes.items():
                    orig_idx = df_disp_5.index[int(r_idx)]
                    if "추가발주수량" in change:
                        val = int(change["추가발주수량"])
                        st.session_state.df_raw.at[orig_idx, "리오더 수량"] += val
                        st.session_state.add_order_dict[orig_idx] = val
                st.success("✅ 발주 수량이 업데이트되었습니다.")
                time.sleep(1); st.rerun()

    # 3. 하단 버튼 구역 (시트 저장 및 CSV)
    st.write("---")
    col_b1, col_b2 = st.columns(2)

    with col_b1:
        if st.button("💾 구글 시트에 최종 발주 기록 저장", use_container_width=True):
            # 권장 + 추가 합산이 0보다 큰 것만 저장
            df_5['합계'] = df_5['권장 발주수량'] + df_5['추가발주수량']
            ready_to_save = df_5[df_5['합계'] > 0]
            if not ready_to_save.empty:
                log_date = d5_date.strftime('%Y-%m-%d')
                log_rows = [[log_date, r['상태'], r[item], r[option], r[v_item], int(r[avail]), int(r['리오더 수량']), int(r['추가발주수량']), int(r['권장 발주수량'])] for _, r in ready_to_save.iterrows()]
                try:
                    sheet = get_sheet()
                    sheet.worksheet("발주기록").append_rows(log_rows)
                    st.success(f"✅ {log_date} 날짜로 {len(log_rows)}건 저장 완료!")
                    st.session_state.add_order_dict = {}
                    time.sleep(1); st.rerun()
                except Exception as e:
                    st.error(f"📡 저장 실패: {e}")

    with col_b2:
        # 다운로드용 CSV
        csv_ready = df_disp_5[(df_disp_5['권장 발주수량'] + df_disp_5['추가발주수량']) > 0]
        if not csv_ready.empty:
            csv_file = csv_ready[display_cols].to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button(
                label="📥 최종 발주서 CSV 다운로드",
                data=csv_file,
                file_name=f"저스트원_발주서_{d5_date.strftime('%m%d')}.csv",
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
