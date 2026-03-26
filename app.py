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



# --- [5단계: 최종 발주 및 엑셀 다운로드] ---
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📋 5단계: 최종 발주 리스트 요약")

    # 1. 설정 및 변수 준비
    p = st.session_state.p
    avail, t7day, t3day = p['av'], p['t7'], p['t3']
    item, option, v_item = p['it'], p['op'], p['vi']
    lt, ss = p['lt'], p['ss']

    df_5 = st.session_state.df_raw.copy()
    
    # 데이터 타입 강제 변환 (안전장치)
    for c in [avail, '리오더 수량', t7day, t3day]:
        if c in df_5.columns:
            df_5[c] = pd.to_numeric(df_5[c], errors='coerce').fillna(0).astype(int)

    # 2. 상단 필터 영역 (사장님 요청 순서 & 긴급 우선)
    f_c1, f_c2, f_c3 = st.columns([1, 2, 1])
    with f_c1:
        # index=1 설정으로 '🚨 긴급'이 기본 선택됨
        m5_filter = st.selectbox("🚦 상태 필터", ["전체보기", "🚨 긴급", "⚠️ 주의", "✅ 정상"], index=1, key="v5_stat_final")
    with f_c2:
        s5_search = st.text_input("🔍 상품명/옵션 검색", placeholder="검색어를 입력하세요...", key="v5_name_final")
    with f_c3:
        d5_date = st.date_input("🗓️ 기록 날짜", datetime.now().date(), key="v5_date_final")

    # 3. 계산 로직 (사장님 원본 공식 유지)
    df_5['일판매량'] = df_5.apply(lambda x: round(x[t7day] / 7) if x[t7day] > 0 else round(x[t3day] / 3), axis=1).astype(int)
    # 권장 발주량 = (일판매량 * (리드타임 + 안전재고)) - (가용재고 + 리오더 수량)
    df_5['권장 발주수량'] = ((df_5['일판매량'] * (lt + ss)) - (df_5[avail] + df_5['리오더 수량'])).clip(lower=0).astype(int)
    
    # 추가발주수량 세션 연동
    if 'add_order_dict' not in st.session_state: st.session_state.add_order_dict = {}
    df_5['추가발주수량'] = df_5.index.map(st.session_state.add_order_dict).fillna(0).astype(int)

    # 상태 판별 함수
    def get_stat_v5_final(r):
        tot = r[avail] + r['리오더 수량']
        day = r['일판매량']
        if day > 0:
            if tot < (day * 3): return "🚨 긴급"
            if tot < (day * 5): return "⚠️ 주의"
        return "✅ 정상"
    df_5['상태'] = df_5.apply(get_stat_v5_final, axis=1)

    # 화면 표시용 가공
    df_disp_5 = df_5.rename(columns={item: "상품명", option: "옵션", v_item: "공급쳐상품명", avail: "가용재고", "리오더 수량": "리오더수량"})
    display_cols = ["상태", "상품명", "옵션", "공급쳐상품명", "가용재고", "리오더수량", "추가발주수량", "권장 발주수량"]

    # 4. 필터링 적용 (실시간 반영)
    if m5_filter != "전체보기":
        df_disp_5 = df_disp_5[df_disp_5["상태"] == m5_filter]
    if s5_search:
        # 상품명과 옵션 둘 다 검색 가능하도록 개선
        df_disp_5 = df_disp_5[df_disp_5["상품명"].astype(str).str.contains(s5_search, case=False) | 
                              df_disp_5["옵션"].astype(str).str.contains(s5_search, case=False)]

    # 5. 데이터 에디터 및 수량 확정
    with st.form("final_order_form_v5_total"):
        if not df_disp_5.empty:
            edited_df = st.data_editor(
                df_disp_5[display_cols],
                use_container_width=True, hide_index=True, key="v5_editor_total",
                column_config={
                    "상태": st.column_config.TextColumn(width="small"),
                    "추가발주수량": st.column_config.NumberColumn("추가발주수량", format="%d"),
                    "권장 발주수량": st.column_config.NumberColumn(format="%d", disabled=True)
                }
            )
        else:
            st.info(f"💡 현재 {m5_filter} 조건에 해당되는 상품이 없습니다.")

        if st.form_submit_button("✅ 수량 확정 및 리오더 반영", use_container_width=True, type="primary"):
            changes = st.session_state["v5_editor_total"].get("edited_rows", {})
            if changes:
                for r_idx, change in changes.items():
                    orig_idx = df_disp_5.index[int(r_idx)]
                    if "추가발주수량" in change:
                        val = int(change["추가발주수량"])
                        st.session_state.df_raw.at[orig_idx, "리오더 수량"] += val
                        st.session_state.add_order_dict[orig_idx] = val
                st.success("✅ 발주 수량이 업데이트되었습니다.")
                time.sleep(1); st.rerun()

    # 6. 하단 버튼 구역 (시트 저장 및 CSV 다운로드)
    st.write("---")
    col_b1, col_b2 = st.columns(2)

    with col_b1:
        # 구글 시트 저장
        if st.button("💾 구글 시트에 최종 발주 기록 저장", use_container_width=True):
            df_5['합계'] = df_5['권장 발주수량'] + df_5['추가발주수량']
            ready_to_save = df_5[df_5['합계'] > 0]
            if not ready_to_save.empty:
                log_date = d5_date.strftime('%Y-%m-%d')
                log_rows = [[log_date, r['상태'], r[item], r[option], r[v_item], int(r[avail]), int(r['리오더 수량']), int(r['추가발주수량']), int(r['권장 발주수량'])] for _, r in ready_to_save.iterrows()]
                try:
                    sheet = get_sheet() # 상단에 정의된 함수 호출
                    sheet.worksheet("발주기록").append_rows(log_rows)
                    st.success(f"✅ {log_date} 날짜로 {len(log_rows)}건 저장 완료!")
                    st.session_state.add_order_dict = {}
                    time.sleep(1); st.rerun()
                except Exception as e:
                    st.error(f"📡 저장 실패: {e}")
            else:
                st.warning("⚠️ 저장할 발주 데이터가 없습니다. (수량이 0보다 커야 함)")

    with col_b2:
        # CSV 다운로드: 필터와 상관없이 발주가 필요한 '전체 데이터' 대상
        csv_target = df_5[(df_5['권장 발주수량'] + df_5['추가발주수량']) > 0]
        if not csv_target.empty:
            csv_disp = csv_target.rename(columns={item: "상품명", option: "옵션", v_item: "공급쳐상품명", avail: "가용재고", "리오더 수량": "리오더수량"})
            csv_file = csv_disp[display_cols].to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button(
                label="📥 최종 발주서 CSV 다운로드",
                data=csv_file,
                file_name=f"저스트원_발주서_{d5_date.strftime('%m%d')}.csv",
                mime="text/csv",
                use_container_width=True
            )
        else:
            st.button("📥 다운로드할 데이터 없음", disabled=True, use_container_width=True)
        
            
# --- [6단계: 전체 히스토리 내역 - 기간 필터 추가] ---
st.divider()
st.subheader("📜 6단계: 전체 히스토리 내역")

try:
    sheet = get_sheet()
    worksheet = sheet.worksheet("발주기록")
    
    # 데이터 가져오기 및 중복 컬럼 방어
    all_values = worksheet.get_all_values()
    
    if len(all_values) > 1:
        headers = all_values[0]
        data = all_values[1:]
        
        # 컬럼명 정제 (빈칸/중복 방지)
        clean_headers = []
        for i, h in enumerate(headers):
            h = h.strip()
            if h == "" or h in clean_headers:
                clean_headers.append(f"empty_{i}")
            else:
                clean_headers.append(h)
        
        df_hist = pd.DataFrame(data, columns=clean_headers)
        df_hist = df_hist[[c for c in df_hist.columns if not c.startswith("empty_")]]

        # [핵심] 날짜 컬럼 형식 변환 (문자열 -> 날짜타입)
        if "날짜" in df_hist.columns:
            df_hist["날짜"] = pd.to_datetime(df_hist["날짜"], errors='coerce').dt.date
            df_hist = df_hist.sort_values(by="날짜", ascending=False)
        
        # --- [상단 필터 UI: 날짜 범위 / 상태 / 검색] ---
        f_c1, f_c2, f_c3 = st.columns([1.5, 1, 1.5])
        
        with f_c1:
            # 1. 달력 필터 (기본값: 최근 7일)
            today = datetime.now().date()
            ten_days_ago = today - timedelta(days=10)
            date_range = st.date_input(
                "🗓️ 조회 기간 선택",
                value=(ten_days_ago, today),
                key="hist_date_range"
            )
        
        with f_c2:
            # 2. 상태 필터
            h_filter = st.selectbox("🚦 상태 필터", ["전체", "🚨 긴급", "⚠️ 주의", "✅ 정상"], key="hist_status_f")
            
        with f_c3:
            # 3. 상품명 검색
            h_search = st.text_input("🔍 상품명/옵션 검색", placeholder="검색어 입력...", key="hist_search_f")
        
        # --- [데이터 필터링 적용] ---
        # A. 날짜 범위 필터
        if len(date_range) == 2:
            start_date, end_date = date_range
            df_hist = df_hist[(df_hist["날짜"] >= start_date) & (df_hist["날짜"] <= end_date)]
            
        # B. 상태 필터
        if h_filter != "전체":
            df_hist = df_hist[df_hist["상태"] == h_filter]
            
        # C. 검색어 필터
        if h_search:
            df_hist = df_hist[
                df_hist["상품명"].astype(str).str.contains(h_search, case=False) |
                df_hist["옵션"].astype(str).str.contains(h_search, case=False)
            ]
            
        # 결과 출력
        if not df_hist.empty:
            st.write(f"✅ 총 **{len(df_hist)}**건의 기록이 검색되었습니다.")
            st.dataframe(df_hist, use_container_width=True, hide_index=True)
            
            # 검색 결과 다운로드 버튼 (선택 사항)
            csv_hist = df_hist.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button("📥 검색 결과 CSV 다운로드", data=csv_hist, file_name=f"발주기록_조회.csv", use_container_width=True)
        else:
            st.warning("🧐 해당 조건에 맞는 기록이 없습니다. 기간이나 필터를 조절해 보세요.")
            
    else:
        st.info("💡 아직 저장된 발주 기록이 없습니다. 5단계에서 저장을 먼저 진행해 주세요.")

except Exception as e:
    st.error(f"📡 시트 로딩 오류: {e}")
    st.info("팁: 구글 시트 '발주기록' 탭 1행에 중복된 제목이나 빈 칸이 없는지 다시 확인해 주세요.")
