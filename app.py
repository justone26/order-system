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


# --- 시트 연결 테스트 모드 ---
st.sidebar.subheader("🔍 시트 연결 상태 점검")
if st.sidebar.button("연결된 시트 탭 목록 확인하기"):
    try:
        sheet = get_sheet() # 기존에 만드신 시트 가져오는 함수
        # 현재 구글 시트 파일 안에 있는 모든 탭(Worksheet) 이름을 가져옵니다.
        worksheets = sheet.worksheets()
        sheet_names = [s.title for s in worksheets]
        
        st.sidebar.success("✅ 시트 연결 성공!")
        st.sidebar.write("**현재 발견된 탭 목록:**")
        for name in sheet_names:
            st.sidebar.code(name) # 탭 이름을 복사하기 좋게 코드로 출력
            
        # 필수 탭이 있는지 자동 체크
        required = ["재고현황", "입고이력", "발주기록"]
        for req in required:
            if req in sheet_names:
                st.sidebar.write(f"✔️ `{req}`: 확인됨")
            else:
                st.sidebar.error(f"❌ `{req}`: 탭을 찾을 수 없습니다!")
                
    except Exception as e:
        st.sidebar.error(f"❌ 시트 연결 실패: {e}")



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
        df = pd.read_csv(up_file) if up_file.name.endswith('.csv') else pd.read_excel(up_file)
        df.columns = df.columns.str.strip()
        
        # [중요] 업로드한 엑셀에 리오더 수량이 없더라도 일단 0으로 세팅 (나중에 시트값으로 덮어씀)
        if "리오더 수량" not in df.columns: 
            df["리오더 수량"] = 0
            
        df = df.fillna("") 
        st.session_state.df_raw = df

    if st.session_state.get('df_raw') is not None:
        cols = st.session_state.df_raw.columns.tolist()
        
        def auto_idx(keys, exclude_keys=None):
            for i, c in enumerate(cols):
                column_name = str(c)
                if exclude_keys and any(ek in column_name for ek in exclude_keys):
                    continue
                if any(k in column_name for k in keys): 
                    return i
            return 0

        c1, c2, c3 = st.columns(3)
        with c1:
            so = st.selectbox("품절 여부", cols, index=auto_idx(['품절']), key="sel_so")
            vn = st.selectbox("공급처", cols, index=auto_idx(['공급처']), key="sel_vn")
            vi = st.selectbox("공급처 상품명", cols, index=auto_idx(['공급처상품명']), key="sel_vi")
        with c2:
            it = st.selectbox("상품명", cols, index=auto_idx(['상품명']), key="sel_it")
            op = st.selectbox("옵션", cols, index=auto_idx(['옵션']), key="sel_op")
            stk = st.selectbox("정상재고", cols, index=auto_idx(['정상재고']), key="sel_stk")
        with c3:
            av = st.selectbox("가용재고", cols, index=auto_idx(['가용재고']), key="sel_av")
            t3 = st.selectbox("3일 판매", cols, index=auto_idx(['3일', '발주'], exclude_keys=['품절']), key="sel_t3")
            t7 = st.selectbox("7일 판매", cols, index=auto_idx(['7일', '1주', '발주'], exclude_keys=['품절']), key="sel_t7")
        
        lt_val = st.number_input("⏳ 리드타임 (일)", value=7, key="inp_lt")
        ss_val = st.number_input("🛡️ 안전재고 (일)", value=3, key="inp_ss")

        if st.button("🚀 데이터 분석 시작", use_container_width=True, type="primary"):
            # 분석용 파라미터 저장
            st.session_state.p = {
                'so': so, 'vn': vn, 'vi': vi, 'it': it, 'op': op, 
                'st': stk, 'av': av, 't3': t3, 't7': t7, 
                'lt': lt_val, 'ss': ss_val
            }
            
            # 데이터 타입 보정
            df_final = st.session_state.df_raw.copy()
            df_final[so] = df_final[so].astype(str).str.strip()
            for num_col in [stk, av, t3, t7]:
                df_final[num_col] = pd.to_numeric(df_final[num_col], errors='coerce').fillna(0).astype(int)

            # ------------------------------------------------------
            # [핵심] 구글 시트('시트1')에서 기존 리오더 수량 가져오기
            # ------------------------------------------------------
            with st.spinner("🔄 구글 시트에서 기존 리오더 수량을 불러오는 중..."):
                try:
                    sheet = get_sheet()
                    m_sh = sheet.worksheet("시트1")
                    master_df = pd.DataFrame(m_sh.get_all_records())

                    if not master_df.empty and "리오더 수량" in master_df.columns:
                        # 엑셀 데이터에서 '리오더 수량' 컬럼이 있으면 제거 (시트 데이터와 중복 방지)
                        if "리오더 수량" in df_final.columns:
                            df_final = df_final.drop(columns=["리오더 수량"])
                        
                        # 상품명(it)과 옵션(op)을 기준으로 병합
                        reorder_data = master_df[[it, op, "리오더 수량"]].copy()
                        df_final = pd.merge(df_final, reorder_data, on=[it, op], how="left")
                        df_final["리오더 수량"] = pd.to_numeric(df_final["리오더 수량"], errors='coerce').fillna(0).astype(int)
                        st.info("✅ 구글 시트의 최신 리오더 수량을 불러왔습니다.")
                    else:
                        if "리오더 수량" not in df_final.columns: df_final["리오더 수량"] = 0
                        st.warning("⚠️ 시트에 리오더 정보가 없어 0으로 시작합니다.")
                except Exception as e:
                    st.error(f"⚠️ 시트 연동 에러: {e}")
                    if "리오더 수량" not in df_final.columns: df_final["리오더 수량"] = 0
            # ------------------------------------------------------

            st.session_state.df_raw = df_final
            st.session_state.analyzed = True
            st.success("데이터 분석 완료! 4단계로 이동하세요.")
            st.rerun()


# ==========================================================
# --- [4단계: 데이터 편집 및 재고 관리 (입고 차감/시트1 연동)] ---
# ==========================================================
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📊 4단계: 데이터 편집 및 재고 관리")

    p = st.session_state.p
    sold_out_col, item, option = p['so'], p['it'], p['op']
    vendor, v_item = p['vn'], p['vi']
    stock, avail, t3day, t7day = p['st'], p['av'], p['t3'], p['t7']
    lt, ss = p['lt'], p['ss']

    df_work = st.session_state.df_raw.copy()
    
    # 데이터 타입 안전장치
    df_work[sold_out_col] = df_work[sold_out_col].astype(str).str.strip()
    for c in [stock, avail, t7day, t3day]:
        df_work[c] = pd.to_numeric(df_work[c], errors='coerce').fillna(0).astype(int)
    if "리오더 수량" not in df_work.columns: df_work["리오더 수량"] = 0
    df_work["리오더 수량"] = pd.to_numeric(df_work["리오더 수량"], errors='coerce').fillna(0).astype(int)
    df_work["리오더 입고수량"] = 0 

    # 1. UI 배치 (필터 및 날짜)
    f_c1, f_c2, f_c3 = st.columns([1, 2, 1])
    filter_m = f_c1.selectbox("🚦 상태 필터", ["전체보기", "정상만", "품절만"], index=1, key="v4_full_filter")
    search_q = f_c2.text_input("🔍 상품명/옵션 검색", placeholder="검색어를 입력하세요...", key="v4_full_search")
    hist_date_4 = f_c3.date_input("🗓️ 입고 날짜", datetime.now(KST).date(), key="v4_full_date")

    # 2. [입고 이력 불러오기] - 시트 이름: '입고기록'
    def get_incoming_sum():
        try:
            sh_h = get_sheet().worksheet("입고기록")
            h_df = pd.DataFrame(sh_h.get_all_records())
            if not h_df.empty:
                return h_df.groupby([item, option])['입고수량'].sum().reset_index()
            return pd.DataFrame(columns=[item, option, "입고수량"])
        except: return pd.DataFrame(columns=[item, option, "입고수량"])

    in_sum_df = get_incoming_sum()
    df_work = pd.merge(df_work, in_sum_df.rename(columns={"입고수량":"과거리오더 입고"}), on=[item, option], how="left").fillna(0)

    # 3. 지표 계산
    df_work['일판매량'] = df_work.apply(lambda x: round(x[t7day] / 7) if x[t7day] > 0 else (round(x[t3day] / 3) if x[t3day] > 0 else 0), axis=1).astype(int)
    df_work['3일 발주수량'] = (df_work['일판매량'] * 3).astype(int)
    df_work['권장발주량'] = ((df_work['일판매량'] * (lt + ss)) - (df_work[avail] + df_work['리오더 수량'])).clip(lower=0).astype(int)

    # 4. 필터 로직
    is_soldout = df_work[sold_out_col].str.contains('품절', na=False)
    if filter_m == "정상만": df_filtered = df_work[~is_soldout]
    elif filter_m == "품절만": df_filtered = df_work[is_soldout]
    else: df_filtered = df_work

    if search_q:
        df_filtered = df_filtered[df_filtered[item].astype(str).str.contains(search_q, case=False) | 
                                  df_filtered[option].astype(str).str.contains(search_q, case=False)]

    # 5. 결과 출력 (컬럼명 변경 및 에디터)
    df_display = df_filtered.rename(columns={sold_out_col: "품절상태", vendor: "공급쳐", v_item: "공급쳐 상품명", item: "상품명", option: "옵션", stock: "정상재고", avail: "가용재고"})
    final_cols = ["품절상태", "공급쳐", "상품명", "옵션", "공급쳐 상품명", "정상재고", "가용재고", "리오더 수량", "리오더 입고수량", "과거리오더 입고", "3일 발주수량", "일판매량", "권장발주량"]

    with st.form("v4_master_full_form"):
        edited_v4 = st.data_editor(
            df_display[final_cols], use_container_width=True, hide_index=True, key="v4_editor_full",
            column_config={
                "과거리오더 입고": st.column_config.NumberColumn(disabled=True, format="%d"),
                "3일 발주수량": st.column_config.NumberColumn(disabled=True, format="%d"),
                "일판매량": st.column_config.NumberColumn(disabled=True, format="%d"),
                "권장발주량": st.column_config.NumberColumn(disabled=True, format="%d"),
                "리오더 입고수량": st.column_config.NumberColumn(format="%d", min_value=0)
            }
        )

        if st.form_submit_button("💾 데이터 저장 및 입고 반영 (차감)", use_container_width=True, type="primary"):
            edits = st.session_state["v4_editor_full"].get("edited_rows", {})
            if edits:
                sheet = get_sheet()
                m_sh = sheet.worksheet("시트1") 
                h_sh = sheet.worksheet("입고기록")
                now_kst = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')

                for r_idx, val in edits.items():
                    if "리오더 입고수량" in val:
                        in_qty = int(val["리오더 입고수량"])
                        orig_idx = df_display.index[int(r_idx)]
                        # 차감 로직
                        new_val = max(0, int(df_display.at[orig_idx, "리오더 수량"]) - in_qty)
                        st.session_state.df_raw.at[orig_idx, "리오더 수량"] = new_val
                        # 이력 저장
                        h_sh.append_row([now_kst, df_display.at[orig_idx, "상품명"], df_display.at[orig_idx, "옵션"], in_qty])
                
                # 마스터 업데이트
                m_sh.update([st.session_state.df_raw.columns.values.tolist()] + st.session_state.df_raw.values.tolist())
                st.success("✅ 입고 차감 및 '시트1' 저장이 완료되었습니다."); time.sleep(1); st.rerun()


# ==========================================================
# --- [5단계: 최종 발주 (합산 및 시트1/발주기록 연동)] ---
# ==========================================================
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📋 5단계: 최종 발주 리스트 요약")

    p = st.session_state.p
    avail, t7day, t3day = p['av'], p['t7'], p['t3']
    item, option, v_item = p['it'], p['op'], p['vi']
    lt, ss = p['lt'], p['ss']

    df_5 = st.session_state.df_raw.copy()
    for c in [avail, t7day, t3day]:
        df_5[c] = pd.to_numeric(df_5[c], errors='coerce').fillna(0).astype(int)
    df_5['리오더 수량'] = pd.to_numeric(df_5.get('리오더 수량', 0), errors='coerce').fillna(0).astype(int)

    # 1. 계산 및 위험군 분류
    df_5['일판매량'] = df_5.apply(lambda x: round(x[t7day] / 7) if x[t7day] > 0 else (round(x[t3day] / 3) if x[t3day] > 0 else 0), axis=1).astype(int)
    df_5['권장 발주수량'] = ((df_5['일판매량'] * (lt + ss)) - (df_5[avail] + df_5['리오더 수량'])).clip(lower=0).astype(int)
    
    if 'add_order_dict' not in st.session_state: st.session_state.add_order_dict = {}
    df_5['추가발주수량'] = df_5.index.map(st.session_state.add_order_dict).fillna(0).astype(int)

    def get_simple_stat(r):
        tot = r[avail] + r['리오더 수량']
        return "🚨 위험군" if r['일판매량'] > 0 and tot < (r['일판매량'] * 5) else "✅ 정상"
    df_5['필터상태'] = df_5.apply(get_simple_stat, axis=1)

    # 위험군 우선 정렬
    df_5 = df_5.sort_values(by=['필터상태', item], ascending=[True, True])

    # 2. 상단 필터
    f1, f2, f3 = st.columns([1.5, 2, 1])
    m5_f = f1.selectbox("🚦 상태 필터", ["🚨 위험군", "✅ 정상"], index=0, key="v5_f_full")
    s5_q = f2.text_input("🔍 상품명 검색", placeholder="상품명을 입력하세요...", key="v5_q_full")
    d5_d = f3.date_input("🗓️ 기준 날짜", datetime.now(KST).date(), key="v5_d_full")

    df_disp_5 = df_5[df_5["필터상태"] == m5_f]
    if s5_q: df_disp_5 = df_disp_5[df_disp_5[item].astype(str).str.contains(s5_q, case=False)]

    # 3. 데이터 에디터 (수량 확정 폼)
    display_map = {"필터상태":"상태", item:"상품명", option:"옵션", v_item:"공급쳐상품명", avail:"가용재고", "리오더 수량":"리오더수량", "추가발주수량":"추가발주수량", "권장 발주수량":"권장 발주수량"}
    
    with st.form("v5_master_full_form"):
        df_edit_5 = df_disp_5[list(display_map.keys())].rename(columns=display_map)
        st.data_editor(
            df_edit_5, use_container_width=True, hide_index=True, key="v5_editor_full",
            column_config={
                "상태": st.column_config.TextColumn(width="small"),
                "추가발주수량": st.column_config.NumberColumn(format="%d"),
                "권장 발주수량": st.column_config.NumberColumn(disabled=True, format="%d")
            }
        )
        
        if st.form_submit_button("✅ 수량 확정 및 리오더 합산", use_container_width=True, type="primary"):
            edits_5 = st.session_state["v5_editor_full"].get("edited_rows", {})
            if edits_5:
                sheet = get_sheet()
                m_sh = sheet.worksheet("시트1") 
                for r_idx, val in edits_5.items():
                    if "추가발주수량" in val:
                        add_v = int(val["추가발주수량"])
                        orig_idx = df_disp_5.index[int(r_idx)]
                        st.session_state.df_raw.at[orig_idx, "리오더 수량"] += add_v
                        st.session_state.add_order_dict[orig_idx] = add_v
                
                m_sh.update([st.session_state.df_raw.columns.values.tolist()] + st.session_state.df_raw.values.tolist())
                st.success("✅ 추가발주가 '시트1'에 반영되었습니다."); time.sleep(1); st.rerun()

    # 4. 하단 버튼 (발주기록 저장 및 CSV 다운로드)
    st.write("---")
    col_b1, col_b2 = st.columns(2)
    
    with col_b1:
        if st.button("💾 구글 시트에 최종 발주 기록 저장", use_container_width=True):
            df_5['합계'] = df_5['권장 발주수량'] + df_5['추가발주수량']
            ready = df_5[df_5['합계'] > 0]
            if not ready.empty:
                now_str = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
                log_rows = [[now_str, r[item], r[option], r[v_item], int(r[avail]), int(r['리오더 수량']), int(r['추가발주수량']), int(r['권장 발주수량'])] for _, r in ready.iterrows()]
                get_sheet().worksheet("발주기록").append_rows(log_rows)
                st.success(f"✅ {now_str} '발주기록' 저장 완료!"); time.sleep(1); st.rerun()
            else:
                st.warning("⚠️ 저장할 수량이 없습니다.")

    with col_b2:
        df_5['합계'] = df_5['권장 발주수량'] + df_5['추가발주수량']
        csv_target = df_5[df_5['합계'] > 0]
        if not csv_target.empty:
            csv_df = csv_target[[item, option, v_item, avail, '리오더 수량', '추가발주수량', '권장 발주수량']].rename(columns={item:"상품명", option:"옵션", v_item:"공급쳐상품명", avail:"가용재고"})
            csv_file = csv_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button("📥 최종 발주서 CSV 다운로드", csv_file, f"발주서_{d5_d.strftime('%m%d')}.csv", "text/csv", use_container_width=True)
        else:
            st.button("📥 다운로드할 데이터 없음", disabled=True, use_container_width=True)

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
