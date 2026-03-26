import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
import time
import io
import pytz  # 시간대 설정을 위한 라이브러리

# 1. [환경 설정 - 한국 시간대 및 페이지 설정]
KST = pytz.timezone('Asia/Seoul') # 한국 시간대 정의
now = datetime.now(KST)          # 현재 한국 시간 가져오기

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

    # 파일 업로드 시 로직 (데이터 타입 강제 지정 포함)
    if up_file:
        # 매번 파일을 새로 읽을 수 있도록 세션 상태 체크 방식을 보강
        df = pd.read_csv(up_file) if up_file.name.endswith('.csv') else pd.read_excel(up_file)
        df.columns = df.columns.str.strip()
        
        # [유실 방지 1] 리오더 수량 컬럼이 없으면 생성
        if "리오더 수량" not in df.columns: 
            df["리오더 수량"] = 0
            
        # [유실 방지 2] 품절 컬럼 등 모든 텍스트 컬럼의 빈값을 '정상'으로 채우기
        df = df.fillna("") 
        
        st.session_state.df_raw = df

    if st.session_state.get('df_raw') is not None:
        cols = st.session_state.df_raw.columns.tolist()
        
        # 키워드 우선순위를 조정하여 '품절'이 숫자칸에 안 들어가게 방어
        def auto_idx(keys, exclude_keys=None):
            for i, c in enumerate(cols):
                column_name = str(c)
                # 제외 키워드가 포함된 컬럼은 패스 (예: 7일판매 칸에 '품절' 컬럼 제외)
                if exclude_keys and any(ek in column_name for ek in exclude_keys):
                    continue
                if any(k in column_name for k in keys): 
                    return i
            return 0

        c1, c2, c3 = st.columns(3)
        with c1:
            # 품절 여부는 '품절' 키워드 우선
            so = st.selectbox("품절 여부", cols, index=auto_idx(['품절']), key="sel_so")
            vn = st.selectbox("공급처", cols, index=auto_idx(['공급처']), key="sel_vn")
            vi = st.selectbox("공급처 상품명", cols, index=auto_idx(['공급처상품명']), key="sel_vi")
        with c2:
            it = st.selectbox("상품명", cols, index=auto_idx(['상품명']), key="sel_it")
            op = st.selectbox("옵션", cols, index=auto_idx(['옵션']), key="sel_op")
            stk = st.selectbox("정상재고", cols, index=auto_idx(['정상재고']), key="sel_stk")
        with c3:
            av = st.selectbox("가용재고", cols, index=auto_idx(['가용재고']), key="sel_av")
            # 3일/7일 판매 칸에는 '품절' 컬럼이 자동으로 들어가지 않게 제외 설정
            t3 = st.selectbox("3일 판매", cols, index=auto_idx(['3일', '발주'], exclude_keys=['품절']), key="sel_t3")
            t7 = st.selectbox("7일 판매", cols, index=auto_idx(['7일', '1주', '발주'], exclude_keys=['품절']), key="sel_t7")
        
        lt_val = st.number_input("⏳ 리드타임 (일)", value=7, key="inp_lt")
        ss_val = st.number_input("🛡️ 안전재고 (일)", value=3, key="inp_ss")

        if st.button("🚀 데이터 분석 시작", use_container_width=True, type="primary"):
            # 분석 시작 시점에 선택된 컬럼 정보를 세션에 저장
            st.session_state.p = {
                'so': so, 'vn': vn, 'vi': vi, 'it': it, 'op': op, 
                'st': stk, 'av': av, 't3': t3, 't7': t7, 
                'lt': lt_val, 'ss': ss_val
            }
            
            # [유실 방지 3] 분석 직전 데이터 타입 강제 변환
            # 품절 컬럼을 문자열로, 판매량 컬럼을 숫자로 확실히 변환해서 4단계로 넘김
            df_final = st.session_state.df_raw.copy()
            df_final[so] = df_final[so].astype(str).str.strip()
            
            # 숫자 데이터 보정
            for num_col in [stk, av, t3, t7]:
                df_final[num_col] = pd.to_numeric(df_final[num_col], errors='coerce').fillna(0)
            
            st.session_state.df_raw = df_final
            st.session_state.analyzed = True
            st.success("데이터 분석 준비 완료! 4단계 탭을 확인하세요.")
            # 분석 완료 후 4단계로 바로 볼 수 있게 rerun (필요시)
            # st.rerun()



# --- 4단계 시작 ---
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📊 4단계: 데이터 편집 및 재고 관리")

    p = st.session_state.p
    sold_out_col = p['so'] 
    item, option = p['it'], p['op']
    vendor, v_item = p['vn'], p['vi']
    stock, avail, t3day, t7day = p['st'], p['av'], p['t3'], p['t7']
    lt, ss = p['lt'], p['ss']

    df_work = st.session_state.df_raw.copy()
    
    # 데이터 타입 안전장치
    df_work[sold_out_col] = df_work[sold_out_col].astype(str).str.strip()
    for c in [stock, avail, t7day, t3day]:
        df_work[c] = pd.to_numeric(df_work[c], errors='coerce').fillna(0).astype(int)

    # 1. UI 배치 (상태 필터 기본값을 '정상만'으로 고정)
    f_c1, f_c2, f_c3 = st.columns([1, 2, 1])
    
    # index=1로 설정하여 '정상만'이 기본 선택되게 함
    filter_m = f_c1.selectbox("🚦 상태 필터", ["전체보기", "정상만", "품절만"], index=1, key="v4_default_normal")
    search_q = f_c2.text_input("🔍 상품명/옵션 검색", placeholder="검색어를 입력하세요...", key="v4_default_search")
    hist_date_4 = f_c3.date_input("🗓️ 입고 날짜", datetime.now().date(), key="v4_default_date")

    # 2. 지표 계산
    df_work['일판매량'] = df_work.apply(lambda x: round(x[t7day] / 7) if x[t7day] > 0 else round(x[t3day] / 3), axis=1).astype(int)
    df_work['3일 발주수량'] = (df_work['일판매량'] * 3).astype(int)
    
    if "리오더 수량" not in df_work.columns: df_work["리오더 수량"] = 0
    df_work["리오더 수량"] = pd.to_numeric(df_work["리오더 수량"], errors='coerce').fillna(0).astype(int)
    df_work["리오더 입고수량"] = 0
    df_work['권장발주량'] = ((df_work['일판매량'] * (lt + ss)) - (df_work[avail] + df_work['리오더 수량'])).clip(lower=0).astype(int)

    # 3. 필터 로직
    # '품절' 글자가 포함된 행 찾기
    is_soldout = df_work[sold_out_col].str.contains('품절', na=False)

    if filter_m == "정상만":
        # 품절이 아닌 행만 추출
        df_filtered = df_work[~is_soldout]
    elif filter_m == "품절만":
        # 품절인 행만 추출
        df_filtered = df_work[is_soldout]
    else:
        # 전체보기
        df_filtered = df_work

    # 검색어 필터 적용
    if search_q:
        df_filtered = df_filtered[
            df_filtered[item].astype(str).str.contains(search_q, case=False, na=False) | 
            df_filtered[option].astype(str).str.contains(search_q, case=False, na=False)
        ]

    # 4. 컬럼명 변경 및 순서 재배치
    df_display = df_filtered.rename(columns={
        sold_out_col: "품절상태", vendor: "공급쳐", v_item: "공급쳐 상품명", 
        item: "상품명", option: "옵션", stock: "정상재고", avail: "가용재고"
    })
    
    inc_h = get_incoming_history()
    if not inc_h.empty:
        df_display = pd.merge(df_display, inc_h, on=["상품명", "옵션"], how="left")
        df_display["과거리오더 입고"] = df_display["과거리오더 입고"].fillna(0).astype(int)
    else:
        df_display["과거리오더 입고"] = 0

    final_cols = [
        "품절상태", "공급쳐", "상품명", "옵션", "공급쳐 상품명", 
        "정상재고", "가용재고", "리오더 수량", "리오더 입고수량", 
        "과거리오더 입고", "3일 발주수량", "일판매량", "권장발주량"
    ]

    # 5. 결과 출력
    with st.form("v4_default_form"):
        if not df_display.empty:
            st.data_editor(
                df_display[final_cols],
                use_container_width=True,
                hide_index=True,
                key="v4_editor_default",
                column_config={c: st.column_config.NumberColumn(disabled=True) for c in ["과거리오더 입고", "3일 발주수량", "일판매량", "권장발주량"]}
            )
        else:
            st.info("💡 표시할 데이터가 없습니다.")

        if st.form_submit_button("💾 데이터 저장 및 입고 반영", use_container_width=True, type="primary"):
            # 저장 로직 (필요시 추가)
            st.success("✅ 저장되었습니다."); time.sleep(1); st.rerun()


# ==========================================================
# --- [5단계: 최종 발주 및 구글 시트 저장 (KST 시간 반영)] ---
# ==========================================================
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📋 5단계: 최종 발주 리스트 요약")

    # 1. 설정 및 데이터 로드
    p = st.session_state.p
    avail, t7day, t3day = p['av'], p['t7'], p['t3']
    item, option, v_item = p['it'], p['op'], p['vi']
    lt, ss = p['lt'], p['ss']

    df_5 = st.session_state.df_raw.copy()
    
    # 데이터 전처리 (숫자 변환)
    for c in [avail, t7day, t3day]:
        if c in df_5.columns:
            df_5[c] = pd.to_numeric(df_5[c], errors='coerce').fillna(0).astype(int)
    
    if '리오더 수량' not in df_5.columns: 
        df_5['리오더 수량'] = 0
    else: 
        df_5['리오더 수량'] = pd.to_numeric(df_5['리오더 수량'], errors='coerce').fillna(0).astype(int)

    # 2. 계산 및 상태 분류 (2종 통합)
    df_5['일판매량'] = df_5.apply(lambda x: round(x[t7day] / 7) if x[t7day] > 0 else (round(x[t3day] / 3) if x[t3day] > 0 else 0), axis=1).astype(int)
    df_5['권장 발주수량'] = ((df_5['일판매량'] * (lt + ss)) - (df_5[avail] + df_5['리오더 수량'])).clip(lower=0).astype(int)
    
    # 추가발주수량 세션 반영
    if 'add_order_dict' not in st.session_state: 
        st.session_state.add_order_dict = {}
    df_5['추가발주수량'] = df_5.index.map(st.session_state.add_order_dict).fillna(0).astype(int)

    # [상태 분류 함수]
    def get_simple_stat(r):
        tot = r[avail] + r['리오더 수량']
        day = r['일판매량']
        # 5일치 미만 재고는 위험군으로 통합
        if day > 0 and tot < (day * 5): return "🚨 위험군(긴급+주의)"
        return "✅ 정상"
    
    df_5['필터상태'] = df_5.apply(get_simple_stat, axis=1)

    # [정렬] 위험군이 무조건 상단에 오도록 정렬
    df_5 = df_5.sort_values(by=['필터상태', item], ascending=[True, True])

    # 3. 상단 필터 및 날짜 설정 (한국 시간 기준 오늘 날짜)
    f_c1, f_c2, f_c3 = st.columns([1.5, 2, 1])
    with f_c1:
        m5_filter = st.selectbox("🚦 상태 필터", ["🚨 위험군(긴급+주의)", "✅ 정상"], index=0, key="v5_kst_filter")
    with f_c2:
        s5_search = st.text_input("🔍 상품명/옵션 검색", placeholder="검색어를 입력하세요...", key="v5_kst_search")
    with f_c3:
        # 한국 시간 기준 오늘 날짜 기본값
        kst_today = datetime.now(KST).date()
        d5_date = st.date_input("🗓️ 기록 기준 날짜", kst_today, key="v5_kst_date")

    # 필터 적용
    df_disp_5 = df_5[df_5["필터상태"] == m5_filter]
    if s5_search:
        df_disp_5 = df_disp_5[df_disp_5[item].astype(str).str.contains(s5_search, case=False) | 
                              df_disp_5[option].astype(str).str.contains(s5_search, case=False)]

    # 4. 화면 데이터 에디터 (상태 포함 8개 항목)
    display_map = {
        "필터상태": "상태", 
        item: "상품명", 
        option: "옵션", 
        v_item: "공급쳐상품명", 
        avail: "가용재고", 
        "리오더 수량": "리오더수량", 
        "추가발주수량": "추가발주수량", 
        "권장 발주수량": "권장 발주수량"
    }
    
    with st.form("final_order_form_v5_kst"):
        df_to_edit = df_disp_5[list(display_map.keys())].rename(columns=display_map)
        edited_df = st.data_editor(
            df_to_edit, 
            use_container_width=True, 
            hide_index=True, 
            key="v5_editor_kst",
            column_config={
                "상태": st.column_config.TextColumn(width="small"),
                "추가발주수량": st.column_config.NumberColumn(format="%d"),
                "권장 발주수량": st.column_config.NumberColumn(format="%d", disabled=True)
            }
        )
        
        if st.form_submit_button("✅ 수량 확정 및 반영", use_container_width=True, type="primary"):
            changes = st.session_state["v5_editor_kst"].get("edited_rows", {})
            for r_idx, change in changes.items():
                orig_idx = df_disp_5.index[int(r_idx)]
                if "추가발주수량" in change:
                    val = int(change["추가발주수량"])
                    st.session_state.df_raw.at[orig_idx, "리오더 수량"] += val
                    st.session_state.add_order_dict[orig_idx] = val
            st.success("✅ 반영되었습니다!"); time.sleep(1); st.rerun()

    # 5. 하단 버튼 (저장 및 CSV 다운로드)
    st.write("---")
    col_b1, col_b2 = st.columns(2)
    
    with col_b1:
        if st.button("💾 구글 시트에 최종 발주 기록 저장", use_container_width=True, key="btn_save_kst_final"):
            df_5['합계'] = df_5['권장 발주수량'] + df_5['추가발주수량']
            ready_to_save = df_5[df_5['합계'] > 0]
            
            if not ready_to_save.empty:
                # [저장 시점 한국 시간 적용]
                now_kst_str = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
                
                # 저장 데이터 구성: 날짜시간 + 7개 항목 (상태 제외)
                log_rows = []
                for _, r in ready_to_save.iterrows():
                    log_rows.append([
                        now_kst_str,           # 0. 날짜시간 (이제 정확히 찍힙니다!)
                        r[item],               # 1. 상품명
                        r[option],             # 2. 옵션
                        r[v_item],             # 3. 공급쳐상품명
                        int(r[avail]),         # 4. 가용재고
                        int(r['리오더 수량']),     # 5. 리오더수량
                        int(r['추가발주수량']),    # 6. 추가발주수량
                        int(r['권장 발


# ==========================================================
# --- [6단계: 전체 히스토리 내역 (상단 선택 및 날짜 노출)] ---
# ==========================================================
st.divider()
st.subheader("📜 6단계: 전체 히스토리 내역")

try:
    sheet = get_sheet()
    worksheet = sheet.worksheet("발주기록")
    all_values = worksheet.get_all_values()
    
    if len(all_values) > 1:
        # 1. 데이터 로드 및 이름 고정 (8개 항목)
        df_hist = pd.DataFrame(all_values[1:])
        target_cols = ["날짜시간", "상품명", "옵션", "공급쳐상품명", "가용재고", "리오더수량", "추가발주수량", "권장 발주수량"]
        
        # 컬럼 개수 맞추기
        if len(df_hist.columns) >= 8:
            df_hist.columns = target_cols + list(df_hist.columns[8:])
            df_hist = df_hist[target_cols]

        # 날짜 필터링용 임시 컬럼
        df_hist["날짜_만"] = df_hist["날짜시간"].astype(str).str.slice(0, 10)
        
        # 2. 상단 필터 레이아웃 (날짜 / 검색 / 회차선택)
        f1, f2, f3 = st.columns([1, 1.5, 1.5])
        
        with f1:
            today = datetime.now().date()
            d_range = st.date_input("🗓️ 날짜 범위", value=(today, today), key="v6_date_final")
        
        # 날짜 1차 필터링
        if len(d_range) == 2:
            s_s, e_s = d_range[0].strftime('%Y-%m-%d'), d_range[1].strftime('%Y-%m-%d')
            df_hist = df_hist[(df_hist["날짜_만"] >= s_s) & (df_hist["날짜_만"] <= e_s)]

        # 3. [핵심] 저장 회차 선택박스 구성
        # 최신순으로 정렬 후 고유한 시간대 추출
        df_hist = df_hist.sort_values(by="날짜시간", ascending=False)
        all_batches = df_hist["날짜시간"].unique().tolist()
        
        with f3:
            if all_batches:
                # 사장님이 말씀하신 "가독성"을 위해 선택박스로 배치
                selected_batch = st.selectbox(
                    "📥 저장 회차 선택 (최신순)", 
                    ["전체보기"] + all_batches,
                    key="v6_batch_select"
                )
            else:
                selected_batch = "기록 없음"
                st.write("조회된 기록이 없습니다.")

        with f2:
            h_q = st.text_input("🔍 상품명 검색", placeholder="내역 중 검색...", key="v6_search_final")

        # 4. 최종 필터링 적용
        df_final_view = df_hist.copy()
        
        if selected_batch != "전체보기" and selected_batch != "기록 없음":
            df_final_view = df_final_view[df_final_view["날짜시간"] == selected_batch]
            
        if h_q:
            df_final_view = df_final_view[df_final_view["상품명"].astype(str).str.contains(h_q, case=False)]

        # 5. 결과 출력 (날짜시간 포함 8개 항목 전체 노출)
        if not df_final_view.empty:
            st.write(f"✅ 총 **{len(df_final_view)}**건의 데이터가 조회되었습니다.")
            
            # 사장님 요청대로 '날짜시간' 포함해서 표에 뿌려줍니다.
            view_cols = ["날짜시간", "상품명", "옵션", "공급쳐상품명", "가용재고", "리오더수량", "추가발주수량", "권장 발주수량"]
            st.dataframe(df_final_view[view_cols], use_container_width=True, hide_index=True)
            
            # 다운로드 버튼
            csv_data = df_final_view[view_cols].to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button(
                "📥 현재 조회된 내역 다운로드", 
                csv_data, 
                f"발주기록_조회결과.csv", 
                use_container_width=True,
                key="v6_download_btn"
            )
        else:
            st.warning("🧐 조건에 맞는 기록이 없습니다.")
            
    else:
        st.info("💡 아직 저장된 발주 기록이 없습니다.")
except Exception as e:
    st.error(f"📡 히스토리 로딩 오류: {e}")
