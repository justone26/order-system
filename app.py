import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
import time
import io
import pytz  # 시간대 설정을 위한 라이브러리
import streamlit.components.v1 as components # <-- 1. 이 라이브러리가 꼭 필요합니다!

# --- [세션 상태 초기화] ---
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'analyzed' not in st.session_state: st.session_state.analyzed = False
if 'p' not in st.session_state: st.session_state.p = {}
if 'add_order_dict' not in st.session_state: st.session_state.add_order_dict = {}
if 'upload_key' not in st.session_state: st.session_state.upload_key = 0

# --- [2. 새로고침 방지 경고창 스크립트] ---
# 이 코드가 실행되면 사용자가 F5를 누를 때 "정말 나갈 거냐"고 물어봅니다.
components.html(
    """
    <script>
    window.onbeforeunload = function() {
        return "데이터 분석 중입니다. 새로고침하면 작업 내용이 사라질 수 있습니다.";
    };
    </script>
    """,
    height=0, # 화면에는 안 보이게 높이를 0으로 설정
)

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


# --- 시트 연결 테스트 모드 (필요할 때만 아래 줄들의 #을 지워서 사용하세요) ---
# st.sidebar.subheader("🔍 시트 연결 상태 점검")
# if st.sidebar.button("연결된 시트 탭 목록 확인하기"):
#     try:
#         sheet = get_sheet() # 기존에 만드신 시트 가져오는 함수
#         # 현재 구글 시트 파일 안에 있는 모든 탭(Worksheet) 이름을 가져옵니다.
#         worksheets = sheet.worksheets()
#         sheet_names = [s.title for s in worksheets]
#         
#         st.sidebar.success("✅ 시트 연결 성공!")
#         st.sidebar.write("**현재 발견된 탭 목록:**")
#         for name in sheet_names:
#             st.sidebar.code(name) # 탭 이름을 복사하기 좋게 코드로 출력
#             
#         # 필수 탭이 있는지 자동 체크
#         required = ["재고현황", "입고이력", "발주기록"]
#         for req in required:
#             if req in sheet_names:
#                 st.sidebar.write(f"✔️ `{req}`: 확인됨")
#             else:
#                 st.sidebar.error(f"❌ `{req}`: 탭을 찾을 수 없습니다!")
#                 
#     except Exception as e:
#         st.sidebar.error(f"❌ 시트 연결 실패: {e}")



# --- [세션 상태 초기화] ---
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'analyzed' not in st.session_state: st.session_state.analyzed = False
if 'p' not in st.session_state: st.session_state.p = {}
if 'add_order_dict' not in st.session_state: st.session_state.add_order_dict = {}

# [추가] 파일 업로드 위젯 리셋용 키 (이게 있어야 파일명이 지워집니다)
if 'upload_key' not in st.session_state: st.session_state.upload_key = 0

st.title("📦 저스트원 통합 재고 관리 v4.0")

tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 관리"])

with tab1:
    # --- 1단계: 데이터 업로드 ---
    st.subheader("📁 1단계: 데이터 업로드")
    
    # 파일 업로드 위젯 (초기화 시 파일명 삭제 기능 포함)
    up_file = st.file_uploader(
        "엑셀/CSV 파일 업로드", 
        type=['xlsx', 'xls', 'csv'], 
        key=f"up_file_{st.session_state.upload_key}"
    )
    
    # [🔄 화면 전체 초기화 버튼]
    if st.button("🔄 화면 전체 초기화", use_container_width=True):
        for key in list(st.session_state.keys()):
            if key != "upload_key": 
                del st.session_state[key]
        st.session_state.upload_key += 1
        st.session_state.analyzed = False 
        st.session_state.df_raw = None
        st.query_params.clear() 
        st.rerun()

    # 데이터 로드 로직
    if up_file:
        if st.session_state.get('df_raw') is None:
            try:
                df = pd.read_csv(up_file) if up_file.name.endswith('.csv') else pd.read_excel(up_file)
                df.columns = df.columns.str.strip()
                if "리오더 수량" not in df.columns: df["리오더 수량"] = 0
                df = df.fillna("") 
                st.session_state.df_raw = df
            except Exception as e:
                st.error(f"파일 로드 오류: {e}")

    # --- 2~3단계: 매핑 및 분석 설정 (파일이 업로드된 경우만 표시) ---
    if st.session_state.get('df_raw') is not None:
        st.divider()
        
        # --- 2단계: 매핑 항목 ---
        st.subheader("📋 2단계: 매핑 항목")
        st.info("💡 업로드된 데이터의 컬럼을 매칭해주세요.")
        cols = st.session_state.df_raw.columns.tolist()
        
        def auto_idx(keys, exclude_keys=None):
            for i, c in enumerate(cols):
                column_name = str(c)
                if exclude_keys and any(ek in column_name for ek in exclude_keys): continue
                if any(k in column_name for k in keys): return i
            return 0

        # 매핑 선택 상자 (3열 배치)
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
            
            # [고정] 3일 판매: '3일 발주합계' 최우선
            t3_target = "3일 발주합계"
            t3_idx = cols.index(t3_target) if t3_target in cols else auto_idx(['3일'], exclude_keys=['1주', '7일', '품절'])
            t3 = st.selectbox("3일 판매", cols, index=t3_idx, key="sel_t3")
            
            # [고정] 7일 판매: '1주발주합계' 최우선
            t7_target = "1주발주합계"
            t7_idx = cols.index(t7_target) if t7_target in cols else auto_idx(['7일', '1주'], exclude_keys=['3일', '품절'])
            t7 = st.selectbox("7일 판매", cols, index=t7_idx, key="sel_t7")

        st.write("") # 간격 조절
        
        # --- 3단계: 데이터 분석 설정 ---
        st.subheader("🚀 3단계: 데이터 분석 설정")
        
        # [리드타임/안전재고 한 줄 배치]
        s1, s2 = st.columns(2)
        with s1:
            lt_val = st.number_input("⏳ 리드타임 (일)", value=7, key="inp_lt")
        with s2:
            ss_val = st.number_input("🛡️ 안전재고 (일)", value=3, key="inp_ss")

        # 분석 시작 버튼
        if st.button("📊 데이터 분석 시작", use_container_width=True, type="primary"):
            st.session_state.p = {
                'so': so, 'vn': vn, 'vi': vi, 'it': it, 'op': op, 
                'st': stk, 'av': av, 't3': t3, 't7': t7, 'lt': lt_val, 'ss': ss_val
            }
            
            # 분석 로직 실행
            df_final = st.session_state.df_raw.copy()
            # (이하 기존의 데이터 가공 및 구글 시트 데이터 병합 로직을 그대로 사용하세요)
            
            st.session_state.df_raw = df_final 
            st.session_state.analyzed = True   
            st.rerun()


# ==========================================================
# --- [4단계: 데이터 편집 및 재고 관리 (과거 입고량 열 복구)] ---
# ==========================================================
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📊 4단계: 데이터 편집 및 재고 관리")

    p = st.session_state.p
    sold_out_col, item, option = p['so'], p['it'], p['op']
    vendor, v_item = p['vn'], p['vi']
    stock, avail, t3day, t7day = p['st'], p['av'], p['t3'], p['t7']
    lt, ss = p['lt'], p['ss']

    # 1. 데이터 준비 및 숫자 타입 변환
    df_work = st.session_state.df_raw.copy()
    num_cols = [stock, avail, t3day, t7day]
    for col in num_cols:
        if col in df_work.columns:
            df_work[col] = pd.to_numeric(df_work[col], errors='coerce').fillna(0).astype(int)
    
    if "리오더 수량" not in df_work.columns: 
        df_work["리오더 수량"] = 0
    df_work["리오더 수량"] = pd.to_numeric(df_work["리오더 수량"], errors='coerce').fillna(0).astype(int)
    df_work["리오더 입고수량"] = 0 

    # 2. ⭐ [복구] 입고 이력 합산 (과거리오더 입고 데이터 가져오기)
    def get_incoming_sum():
        try:
            sh_h = get_sheet().worksheet("입고기록")
            h_data = sh_h.get_all_records()
            if h_data:
                h_df = pd.DataFrame(h_data)
                # 상품명과 옵션별로 입고수량 합산
                return h_df.groupby(['상품명', '옵션'])['입고수량'].sum().reset_index()
            return pd.DataFrame(columns=['상품명', '옵션', '입고수량'])
        except: 
            return pd.DataFrame(columns=['상품명', '옵션', '입고수량'])

    in_sum_df = get_incoming_sum()
    # 원본 데이터에 과거 입고 합계 병합
    df_work = pd.merge(df_work, in_sum_df.rename(columns={"입고수량":"과거리오더 입고"}), 
                       left_on=[item, option], right_on=['상품명', '옵션'], how="left").fillna(0)

    # 3. 지표 계산 (정수 반올림)
    def calc_daily_sales_int(row):
        t7, t3 = row[t7day], row[t3day]
        if t7 > 0: return int(round(t7 / 7))
        elif t3 > 0: return int(round(t3 / 3))
        return 0

    df_work['일판매'] = df_work.apply(calc_daily_sales_int, axis=1)
    df_work['3일발주'] = (df_work['일판매'] * 3).astype(int)
    df_work['권장발주'] = ((df_work['일판매'] * (lt + ss)) - (df_work[avail] + df_work['리오더 수량'])).clip(lower=0).astype(int)

    # 4. 상단 레이아웃
    f_c1, f_c2, f_c3 = st.columns([1, 2, 1])
    with f_c1: filter_m = st.selectbox("🚦 필터", ["전체보기", "정상만", "품절만"], index=1, key="v4_fix_filter")
    with f_c2: search_q = st.text_input("🔍 검색", placeholder="상품명 또는 옵션...", key="v4_fix_search")
    with f_c3: hist_date_4 = st.date_input("🗓️ 입고 날짜", datetime.now(KST).date(), key="v4_fix_date")

    # 필터링
    is_soldout = df_work[sold_out_col].astype(str).str.contains('품절', na=False)
    df_filtered = df_work[~is_soldout] if filter_m == "정상만" else (df_work[is_soldout] if filter_m == "품절만" else df_work)
    if search_q:
        df_filtered = df_filtered[df_filtered[item].astype(str).str.contains(search_q, case=False) | 
                                  df_filtered[option].astype(str).str.contains(search_q, case=False)]

    # 5. 화면 출력 설정 (과거 입고 열 포함)
    df_display = df_filtered.rename(columns={
        sold_out_col: "상태", vendor: "공급쳐", v_item: "공급상품명", 
        item: "상품명", option: "옵션", stock: "정상", avail: "가용"
    })
    
    # 사장님이 찾으시던 '과거리오더 입고'를 리스트에 다시 넣었습니다.
    final_cols = ["상태", "공급쳐", "상품명", "옵션", "공급상품명", "정상", "가용", "리오더 수량", "리오더 입고수량", "과거리오더 입고", "3일발주", "일판매", "권장발주"]

    with st.form("v4_fix_master_form"):
        edited_v4 = st.data_editor(
            df_display[final_cols], 
            use_container_width=True, 
            hide_index=True, 
            key="v4_editor_fix",
            column_config={
                "상태": st.column_config.TextColumn(width=60),
                "공급쳐": st.column_config.TextColumn(width=80),
                "상품명": st.column_config.TextColumn(width=350), # 열이 늘어나서 조금 조정했습니다.
                "옵션": st.column_config.TextColumn(width=100),
                "과거리오더 입고": st.column_config.NumberColumn("과거입고", width=70, format="%d"),
                "리오더 입고수량": st.column_config.NumberColumn("입고입력", width=70, format="%d", min_value=0),
                "정상": st.column_config.NumberColumn(width=50, format="%d"),
                "가용": st.column_config.NumberColumn(width=50, format="%d"),
                "리오더 수량": st.column_config.NumberColumn(width=70, format="%d"),
                "일판매": st.column_config.NumberColumn(width=50, format="%d"),
                "권장발주": st.column_config.NumberColumn(width=70, format="%d"),
            }
        )

        if st.form_submit_button("💾 데이터 저장 및 입고 반영", use_container_width=True, type="primary"):
            user_edits = st.session_state["v4_editor_fix"].get("edited_rows", {})
            if user_edits:
                m_sh, h_sh = get_sheet().worksheet("시트1"), get_sheet().worksheet("입고기록")
                save_time = f"{hist_date_4.strftime('%Y-%m-%d')} {datetime.now(KST).strftime('%H:%M:%S')}"

                for r_idx_str, changes in user_edits.items():
                    target_idx = df_display.index[int(r_idx_str)]
                    if "리오더 수량" in changes:
                        st.session_state.df_raw.at[target_idx, "리오더 수량"] = int(changes["리오더 수량"])
                    if "리오더 입고수량" in changes:
                        in_qty = int(changes["리오더 입고수량"])
                        if in_qty > 0:
                            old_v = int(st.session_state.df_raw.at[target_idx, "리오더 수량"])
                            st.session_state.df_raw.at[target_idx, "리오더 수량"] = max(0, old_v - in_qty)
                            h_sh.append_row([save_time, str(df_display.at[target_idx, "상품명"]), str(df_display.at[target_idx, "옵션"]), in_qty])

                df_to_save = st.session_state.df_raw.copy().fillna("").astype(str)
                m_sh.update([df_to_save.columns.values.tolist()] + df_to_save.values.tolist())
                st.success(f"✅ 저장 및 차감 완료!"); time.sleep(0.5); st.rerun()



# ==========================================================
# --- [5단계: 추가 오더 관리 (2단 필터 및 세트 노출 최적화)] ---
# ==========================================================
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📦 5단계: 추가 오더(위험군) 관리")

    p = st.session_state.p
    sold_out_col, item, option = p['so'], p['it'], p['op']
    vendor_item_col = p['vi']
    avail_col = p['av']
    lt, ss = p['lt'], p['ss']

    # 1. 데이터 타입 변환 및 준비
    df_v5 = st.session_state.df_raw.copy()
    if "리오더 수량" not in df_v5.columns: df_v5["리오더 수량"] = 0
    
    for c in [avail_col, "리오더 수량", p['t7'], p['t3']]:
        if c in df_v5.columns:
            df_v5[c] = pd.to_numeric(df_v5[c], errors='coerce').fillna(0).astype(int)

    # 2. 과거 입고량 합산 데이터 병합
    def get_v5_history():
        try:
            sh_h = get_sheet().worksheet("입고기록")
            h_df = pd.DataFrame(sh_h.get_all_records())
            return h_df.groupby(['상품명', '옵션'])['입고수량'].sum().reset_index()
        except: return pd.DataFrame(columns=['상품명', '옵션', '입고수량'])

    in_sum_v5 = get_v5_history()
    df_v5 = pd.merge(df_v5, in_sum_v5.rename(columns={"입고수량":"과거입고"}), 
                     left_on=[item, option], right_on=['상품명', '옵션'], how="left").fillna(0)

    # 3. 지표 계산 및 위험 아이콘 분류
    df_v5['일판매'] = df_v5.apply(lambda r: int(round(r[p['t7']]/7)) if r[p['t7']]>0 else (int(round(r[p['t3']]/3)) if r[p['t3']]>0 else 0), axis=1)
    df_v5['권장발주'] = ((df_v5['일판매'] * (lt + ss)) - (df_v5[avail_col] + df_v5['리오더 수량'])).clip(lower=0).astype(int)

    def get_icon(row):
        if row['권장발주'] >= 10: return "🚨 고위험"
        elif row['권장발주'] > 0: return "⚠️ 주의"
        return "✅ 정상"
    df_v5['경고'] = df_v5.apply(get_icon, axis=1)

    # 4. 상단 레이아웃 (🚦 2단 필터 / 🔍 검색 / 🗓️ 날짜)
    v5_f1, v5_f2, v5_f3 = st.columns([1, 1.5, 1])
    with v5_f1:
        v5_filter = st.selectbox("🚦 보기 설정", ["🚨 고위험/주의", "✅ 정상"], key="v5_main_filter")
    with v5_f2:
        search_v5 = st.text_input("🔍 검색", placeholder="상품명 입력...", key="v5_search")
    with v5_f3:
        date_v5 = st.date_input("🗓️ 발주 날짜", datetime.now(KST).date(), key="v5_date")

    # 5. [필터 로직] 품절 제외 및 세트 노출 적용
    df_ns = df_v5[~df_v5[sold_out_col].astype(str).str.contains('품절', na=False)].copy()

    if v5_filter == "🚨 고위험/주의":
        # 권장발주가 1개라도 있는 상품의 전체 옵션 추출
        danger_names = df_ns[df_ns['권장발주'] > 0][item].unique()
        df_v5_filtered = df_ns[df_ns[item].isin(danger_names)].copy()
    else:
        # 권장발주가 0인 상품들만 추출 (정상)
        danger_names = df_ns[df_ns['권장발주'] > 0][item].unique()
        df_v5_filtered = df_ns[~df_ns[item].isin(danger_names)].copy()

    # 정렬 및 검색
    df_v5_filtered = df_v5_filtered.sort_values(['권장발주', item, option], ascending=[False, True, True])
    if search_v5:
        df_v5_filtered = df_v5_filtered[df_v5_filtered[item].str.contains(search_v5, case=False)]

    # 6. 화면 출력용 컬럼 정리 (사장님 요청 순서)
    df_v5_display = df_v5_filtered.rename(columns={
        item: "상품명", option: "옵션", vendor_item_col: "공급처상품명", avail_col: "가용"
    })
    df_v5_display["추가오더"] = 0
    final_cols = ["경고", "상품명", "옵션", "공급처상품명", "가용", "리오더 수량", "추가오더", "과거입고", "일판매", "권장발주"]

    # 7. 데이터 에디터 (열 너비 최적화)
    st.info(f"📊 {v5_filter} - {len(df_v5_filtered)}건 노출 중")
    edited_data = st.data_editor(
        df_v5_display[final_cols],
        use_container_width=True,
        hide_index=True,
        key="v5_editor_final",
        column_config={
            "경고": st.column_config.TextColumn(width=80),
            "상품명": st.column_config.TextColumn(width=300),
            "공급처상품명": st.column_config.TextColumn(width=180),
            "추가오더": st.column_config.NumberColumn(width=80, format="%d", min_value=0),
            "가용": st.column_config.NumberColumn(width=60, format="%d"),
            "권장발주": st.column_config.NumberColumn(width=70, format="%d"),
        }
    )

    # 8. 하단 버튼 (나란히 배치)
    btn_c1, btn_c2 = st.columns(2)
    with btn_c1:
        if st.button("📝 발주 기록 저장", use_container_width=True, type="primary"):
            user_edits = st.session_state["v5_editor_final"].get("edited_rows", {})
            if user_edits:
                m_sh, o_sh = get_sheet().worksheet("시트1"), get_sheet().worksheet("발주기록")
                save_time = f"{date_v5.strftime('%Y-%m-%d')} {datetime.now(KST).strftime('%H:%M:%S')}"
                for r_idx, changes in user_edits.items():
                    idx = df_v5_display.index[int(r_idx)]
                    if "추가오더" in changes and changes["추가오더"] > 0:
                        qty = int(changes["추가오더"])
                        st.session_state.df_raw.at[idx, "리오더 수량"] = int(st.session_state.df_raw.at[idx, "리오더 수량"]) + qty
                        o_sh.append_row([save_time, str(df_v5_display.at[idx, "공급처상품명"]), str(df_v5_display.at[idx, "상품명"]), str(df_v5_display.at[idx, "옵션"]), qty])
                
                df_to_save = st.session_state.df_raw.copy().fillna("").astype(str)
                m_sh.update([df_to_save.columns.values.tolist()] + df_to_save.values.tolist())
                st.success("✅ 발주 내역 저장 완료!"); time.sleep(0.5); st.rerun()

    with btn_c2:
        csv_data = df_v5_display[final_cols].to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 엑셀 다운로드", data=csv_data, file_name=f"발주요청_{date_v5}.csv", use_container_width=True)



# ==========================================================
# --- [6단계: 전체 히스토리 관리 (날짜 -> 검색 -> 상품명 -> 회차)] ---
# ==========================================================
if st.session_state.get('analyzed'):
    st.divider()
    st.subheader("📜 6단계: 전체 히스토리 관리")

    # 사장님 요청 순서: 날짜 -> 검색버튼 -> 상품명 -> 회차
    f1, f2, f3, f4 = st.columns([1, 0.5, 1.2, 1.2])
    
    with f1:
        today = datetime.now(KST).date()
        d_range = st.date_input("🗓️ 날짜 범위", value=(today, today), key="v6_date_final")
    
    with f2:
        st.write("") # 높이 맞춤용
        st.write("") 
        # 날짜 바로 옆에 배치된 검색 버튼 (데이터를 불러오는 트리거)
        search_trigger = st.button("🔍 검색", use_container_width=True, type="primary", key="v6_search_btn")

    with f3:
        h_q = st.text_input("🔍 상품명 검색", placeholder="결과 내 검색...", key="v6_search_final")
        
    with f4:
        # 회차 선택박스
        selected_batch = st.selectbox("📥 저장 회차 선택", ["전체보기"], key="v6_batch_select")

    # 검색 버튼 클릭 시 로직
    if search_trigger:
        try:
            with st.spinner("📡 기록을 불러오는 중..."):
                sheet = get_sheet()
                worksheet = sheet.worksheet("발주기록")
                all_values = worksheet.get_all_values()
            
            if len(all_values) > 1:
                df_hist = pd.DataFrame(all_values[1:])
                target_cols = ["날짜시간", "상품명", "옵션", "공급쳐상품명", "가용재고", "리오더수량", "추가발주수량", "권장 발주수량"]
                
                if len(df_hist.columns) >= 8:
                    df_hist.columns = target_cols + list(df_hist.columns[8:])
                    df_hist = df_hist[target_cols]

                # --- 데이터 필터링 시작 ---
                df_hist["날짜_만"] = df_hist["날짜시간"].astype(str).str.slice(0, 10)
                
                # 1. 날짜 필터 (가장 우선 적용)
                if len(d_range) == 2:
                    s_s, e_s = d_range[0].strftime('%Y-%m-%d'), d_range[1].strftime('%Y-%m-%d')
                    df_hist = df_hist[(df_hist["날짜_만"] >= s_s) & (df_hist["날짜_만"] <= e_s)]

                # 2. 결과 정렬 (최신순)
                df_hist = df_hist.sort_values(by="날짜시간", ascending=False)

                # --- 결과 출력 ---
                if not df_hist.empty:
                    # 회차 선택박스 업데이트를 위한 세션 저장 (필요시)
                    st.session_state.v6_data = df_hist
                    
                    # 상품명 검색어 필터링 (결과 내 검색)
                    df_view = df_hist.copy()
                    if h_q:
                        df_view = df_view[df_view["상품명"].astype(str).str.contains(h_q, case=False)]

                    st.write(f"✅ 총 **{len(df_view)}**건의 내역이 조회되었습니다.")
                    st.dataframe(df_view, use_container_width=True, hide_index=True)
                    
                    csv_data = df_view.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
                    st.download_button("📥 내역 다운로드(CSV)", csv_data, "발주기록_검색결과.csv", use_container_width=True)
                else:
                    st.warning("🧐 해당 날짜 범위에 기록이 없습니다.")
            else:
                st.info("💡 아직 저장된 발주 기록이 없습니다.")
        except Exception as e:
            st.error(f"📡 데이터 로딩 오류: {e}")
