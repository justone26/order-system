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



   # --- 2~3단계: 매핑 및 분석 설정 ---
if st.session_state.get('df_raw') is not None:
    st.divider()
    
    # --- 2단계: 매핑 항목 ---
    st.subheader("📋 2단계: 매핑 항목")
    st.info("💡 좌측은 기본 정보, 우측은 수량 및 날짜 정보를 매칭해주세요.")
    cols = st.session_state.df_raw.columns.tolist()
    
    def auto_idx(keys, exclude_keys=None):
        for i, c in enumerate(cols):
            column_name = str(c)
            if exclude_keys and any(ek in column_name for ek in exclude_keys): continue
            if any(k in column_name for k in keys): return i
        return 0

    # 5개씩 2열로 배치 (좌: 기본정보 / 우: 수량 및 날짜)
    c_left, c_right = st.columns(2)
    
    with c_left:
        st.markdown("##### [ 기본 정보 ]")
        it = st.selectbox("📦 상품명", cols, index=auto_idx(['상품명']), key="sel_it")
        op = st.selectbox("🎨 옵션", cols, index=auto_idx(['옵션']), key="sel_op")
        vn = st.selectbox("🏭 공급처", cols, index=auto_idx(['공급처']), key="sel_vn")
        vi = st.selectbox("🆔 공급처 상품명", cols, index=auto_idx(['공급처상품명']), key="sel_vi")
        so = st.selectbox("🚫 품절 여부", cols, index=auto_idx(['품절']), key="sel_so")

    with c_right:
        st.markdown("##### [ 수량 및 날짜 ]")
        av = st.selectbox("✅ 가용재고", cols, index=auto_idx(['가용재고']), key="sel_av")
        stk = st.selectbox("📦 정상재고", cols, index=auto_idx(['정상재고']), key="sel_stk")
        
        # 3일 판매: '3일 발주합계' 최우선
        t3_target = "3일 발주합계"
        t3_idx = cols.index(t3_target) if t3_target in cols else auto_idx(['3일'], exclude_keys=['1주', '7일', '품절'])
        t3 = st.selectbox("🔥 3일 판매", cols, index=t3_idx, key="sel_t3")
        
        # 7일 판매: '1주발주합계' 최우선
        t7_target = "1주발주합계"
        t7_idx = cols.index(t7_target) if t7_target in cols else auto_idx(['7일', '1주'], exclude_keys=['3일', '품절'])
        t7 = st.selectbox("📅 7일 판매", cols, index=t7_idx, key="sel_t7")
        
        # 등록일 추가
        reg = st.selectbox("📆 상품 등록일", cols, index=auto_idx(['등록일', '등록일자', '최초등록']), key="sel_reg")

    st.write("") # 간격 조절
    
    # --- 3단계: 데이터 분석 설정 ---
    st.subheader("🚀 3단계: 데이터 분석 설정")
    
    s1, s2 = st.columns(2)
    with s1:
        lt_val = st.number_input("⏳ 리드타임 (일)", value=7, key="inp_lt")
    with s2:
        ss_val = st.number_input("🛡️ 안전재고 (일)", value=3, key="inp_ss")

    # 분석 시작 버튼
    if st.button("📊 데이터 분석 시작", use_container_width=True, type="primary"):
        st.session_state.p = {
            'so': so, 'vn': vn, 'vi': vi, 'it': it, 'op': op, 
            'st': stk, 'av': av, 't3': t3, 't7': t7, 'reg': reg,
            'lt': lt_val, 'ss': ss_val
        }
        
        df_final = st.session_state.df_raw.copy()
        
        # 등록일 날짜 형식 변환
        if reg in df_final.columns:
            df_final[reg] = pd.to_datetime(df_final[reg], errors='coerce')
        
        st.session_state.df_raw = df_final 
        st.session_state.analyzed = True   
        st.rerun()





# ==========================================================
# --- [4단계: 데이터 편집 및 재고 관리 (등록일 보정 로직 적용)] ---
# ==========================================================
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📊 4단계: 데이터 편집 및 재고 관리")

    p = st.session_state.p
    sold_out_col, item, option = p['so'], p['it'], p['op']
    vendor, v_item = p['vn'], p['vi']
    stock, avail, t3day, t7day = p['st'], p['av'], p['t3'], p['t7']
    reg_date_col = p.get('reg')  # 2단계에서 매핑한 등록일 컬럼
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

    # 2. 입고 이력 합산 (과거리오더 입고 데이터 가져오기)
    @st.cache_data(ttl=60)
    def get_incoming_sum():
        try:
            sh_h = get_sheet().worksheet("입고기록")
            h_data = sh_h.get_all_records()
            if h_data:
                h_df = pd.DataFrame(h_data)
                return h_df.groupby(['상품명', '옵션'])['입고수량'].sum().reset_index()
            return pd.DataFrame(columns=['상품명', '옵션', '입고수량'])
        except: 
            return pd.DataFrame(columns=['상품명', '옵션', '입고수량'])

    in_sum_df = get_incoming_sum()
    df_work = pd.merge(df_work, in_sum_df.rename(columns={"입고수량":"과거리오더 입고"}), 
                       left_on=[item, option], right_on=['상품명', '옵션'], how="left").fillna(0)

    # 3. ⭐ [핵심 수정] 신상품 보정 일판매량 계산 로직
    def calc_daily_sales_with_reg(row):
        t7, t3 = row[t7day], row[t3day]
        
        # 등록일 기준 보정 계산
        if reg_date_col and reg_date_col in row and pd.notnull(row[reg_date_col]):
            # 기준 날짜(오늘)와 등록일의 차이 계산
            today = datetime.now(KST).date()
            reg_dt = row[reg_date_col].date() if hasattr(row[reg_date_col], 'date') else pd.to_datetime(row[reg_date_col]).date()
            days_diff = (today - reg_dt).days
            
            # 등록 3일 이내 신상품인 경우
            if 0 <= days_diff < 3:
                actual_days = days_diff + 1  # 당일은 1일차
                return int(round(t3 / actual_days)) if t3 > 0 else 0
        
        # 일반 상품 (기존 로직)
        if t7 > 0: return int(round(t7 / 7))
        elif t3 > 0: return int(round(t3 / 3))
        return 0

    df_work['일판매'] = df_work.apply(calc_daily_sales_with_reg, axis=1)
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

    # 5. 화면 출력 설정
    df_display = df_filtered.rename(columns={
        sold_out_col: "상태", vendor: "공급쳐", v_item: "공급상품명", 
        item: "상품명", option: "옵션", stock: "정상", avail: "가용"
    })
    
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
                "상품명": st.column_config.TextColumn(width=350),
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
# --- [5단계: 최종 발주 (신상품 보정 및 변동 선별 저장)] ---
# ==========================================================
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📋 5단계: 최종 발주 리스트 요약")

    # [1. 기본 설정]
    p = st.session_state.p
    sold_out_col = p['so']
    avail, t7day, t3day = p['av'], p['t7'], p['t3']
    item, option, v_item = p['it'], p['op'], p['vi']
    reg_date_col = p.get('reg') # 매핑된 등록일 컬럼
    lt, ss = p['lt'], p['ss']

    # 데이터 복사 및 숫자 형변환
    df_v5_base = st.session_state.df_raw.copy()
    for c in [avail, t7day, t3day]:
        if c in df_v5_base.columns:
            df_v5_base[c] = pd.to_numeric(df_v5_base[c], errors='coerce').fillna(0).astype(int)
    
    if "리오더 수량" not in df_v5_base.columns: df_v5_base["리오더 수량"] = 0
    df_v5_base['리오더 수량'] = pd.to_numeric(df_v5_base['리오더 수량'], errors='coerce').fillna(0).astype(int)

    # [2. 과거 입고 데이터 불러오기]
    @st.cache_data(ttl=60)
    def get_v5_history_data():
        try:
            sh_h = get_sheet().worksheet("입고기록")
            h_df = pd.DataFrame(sh_h.get_all_records())
            if not h_df.empty:
                h_df['입고수량'] = pd.to_numeric(h_df['입고수량'], errors='coerce').fillna(0)
                h_df['상품명_match'] = h_df['상품명'].astype(str).str.strip()
                h_df['옵션_match'] = h_df['옵션'].astype(str).str.strip()
                return h_df.groupby(['상품명_match', '옵션_match'])['입고수량'].sum().reset_index()
            return pd.DataFrame(columns=['상품명_match', '옵션_match', '입고수량'])
        except: return pd.DataFrame(columns=['상품명_match', '옵션_match', '입고수량'])

    df_h_data = get_v5_history_data().rename(columns={"입고수량": "과거입고_참고"})

    # [3. 화면 표시용 데이터 구성]
    df_v5_base['_m_i'] = df_v5_base[item].astype(str).str.strip()
    df_v5_base['_m_o'] = df_v5_base[option].astype(str).str.strip()

    df_display = pd.merge(
        df_v5_base, 
        df_h_data, 
        left_on=['_m_i', '_m_o'], 
        right_on=['상품명_match', '옵션_match'], 
        how="left"
    ).fillna({"과거입고_참고": 0})
    
    df_display.index = df_v5_base.index
    df_display = df_display.drop(columns=['_m_i', '_m_o', '상품명_match', '옵션_match'], errors='ignore')

    # [4. 계산 로직 - ⭐ 등록일 기준 신상품 보정 적용]
    def calculate_v5_daily_sales(row, target_date):
        t7, t3 = row[t7day], row[t3day]
        
        # 신상품 판단 (등록일 매핑 정보가 있을 때)
        if reg_date_col and reg_date_col in row and pd.notnull(row[reg_date_col]):
            reg_dt = row[reg_date_col].date() if hasattr(row[reg_date_col], 'date') else pd.to_datetime(row[reg_date_col]).date()
            days_diff = (target_date - reg_dt).days
            
            # 등록 3일 이내 (0, 1, 2일차)
            if 0 <= days_diff < 3:
                actual_days = days_diff + 1
                return round(t3 / actual_days) if t3 > 0 else 0
        
        # 일반 상품 로직
        if t7 > 0: return round(t7 / 7)
        elif t3 > 0: return round(t3 / 3)
        return 0

    # 기준 날짜 선택 UI (계산에 사용됨)
    f1, f2, f3 = st.columns([1.5, 2, 1])
    m5_f = f1.selectbox("🚦 상태 필터", ["🚨 고위험/주의", "✅ 전체정상"], key="v5_filter_fixed")
    s5_q = f2.text_input("🔍 상품명 검색", key="v5_search_fixed")
    d5_d = f3.date_input("🗓️ 기준 날짜", datetime.now(KST).date(), key="v5_date_fixed")

    # 일판매량 및 권장발주 계산 적용
    df_display['일판매량'] = df_display.apply(lambda x: calculate_v5_daily_sales(x, d5_d), axis=1).astype(int)
    df_display['권장 발주수량'] = ((df_display['일판매량'] * (lt + ss)) - (df_display[avail] + df_display['리오더 수량'])).clip(lower=0).astype(int)
    
    if 'add_order_dict' not in st.session_state: st.session_state.add_order_dict = {}
    df_display['추가발주수량'] = df_display.index.map(st.session_state.add_order_dict).fillna(0).astype(int)

    def get_stat(r):
        if r['권장 발주수량'] >= 10: return "🚨 고위험"
        elif r['권장 발주수량'] > 0: return "⚠️ 주의"
        return "✅ 정상"
    df_display['상태분류'] = df_display.apply(get_stat, axis=1)

    # [5. 필터 및 정렬]
    df_ns = df_display[~df_display[sold_out_col].astype(str).str.contains('품절', na=False)].copy()

    if item in df_ns.columns:
        if m5_f == "🚨 고위험/주의":
            danger_names = df_ns[df_ns['권장 발주수량'] > 0][item].unique()
            df_final_view = df_ns[df_ns[item].isin(danger_names)].copy()
        else:
            danger_names = df_ns[df_ns['권장 발주수량'] > 0][item].unique()
            df_final_view = df_ns[~df_ns[item].isin(danger_names)].copy()
        
        df_final_view = df_final_view.sort_values(by=[item, option], ascending=[True, True])
        if s5_q: 
            df_final_view = df_final_view[df_final_view[item].astype(str).str.contains(s5_q, case=False)]

   # [6. 데이터 에디터 출력]
    display_map = {
        "상태분류": "상태", item: "상품명", option: "옵션", v_item: "공급쳐상품명", 
        avail: "가용재고", "리오더 수량": "리오더수량", "추가발주수량": "추가발주수량", 
        "권장 발주수량": "권장 발주수량", "과거입고_참고": "과거입고" 
    }
    
    with st.form("v5_form_fixed"):
        valid_cols = [c for c in display_map.keys() if c in df_final_view.columns]
        df_edit = df_final_view[valid_cols].rename(columns=display_map)
        
        # ⭐ 중요: 에디터 출력 (key값 확인)
        st.data_editor(
            df_edit, use_container_width=True, hide_index=True, key="v5_editor_fixed",
            column_config={
                "상태": st.column_config.TextColumn(width="small"),
                "상품명": st.column_config.TextColumn(width=280),
                "추가발주수량": st.column_config.NumberColumn(format="%d", min_value=0),
            }
        )
        
        if st.form_submit_button("✅ 수량 확정 및 리오더 합산", use_container_width=True, type="primary"):
            # 1. 세션에서 수정된 데이터 가져오기
            edits = st.session_state["v5_editor_fixed"].get("edited_rows", {})
            
            if edits:
                try:
                    with st.spinner("🔄 시트에 반영 중..."):
                        m_sh = get_sheet().worksheet("시트1")
                        
                        # 2. 수정된 내역을 하나씩 순회
                        for r_idx_str, val in edits.items():
                            if "추가발주수량" in val:
                                r_idx = int(r_idx_str)
                                # ⭐ 핵심: 현재 화면(df_final_view)의 줄 번호로 원본 인덱스 추출
                                orig_idx = df_final_view.index[r_idx]
                                
                                input_val = int(val["추가발주수량"])
                                
                                # 원본 데이터(df_raw)에 즉시 합산
                                st.session_state.df_raw.at[orig_idx, "리오더 수량"] += input_val
                                # 세션 딕셔너리에도 기록 (히스토리용)
                                st.session_state.add_order_dict[orig_idx] = input_val
                        
                        # 3. 전체 데이터 시트 업데이트
                        df_to_save = st.session_state.df_raw.copy().fillna("").astype(str)
                        m_sh.update([df_to_save.columns.values.tolist()] + df_to_save.values.tolist())
                        
                        st.success("✅ 리오더 수량이 성공적으로 합산되었습니다!")
                        time.sleep(1)
                        st.rerun()
                except Exception as e:
                    st.error(f"❌ 저장 중 오류 발생: {e}")
            else:
                st.warning("⚠️ 입력된 수량이 없습니다. 숫자를 넣고 '엔터'를 친 후 버튼을 눌러주세요.")

  # ==========================================================
# --- [7. 하단 버튼 - 실질적 변동(입력)이 있는 상품만 기록 저장] ---
# ==========================================================
st.write("---")
col_b1, col_b2 = st.columns(2)

with col_b1:
    # [수정 포인트]: 사장님이 직접 수치를 기입하거나 수정한 항목만 필터링합니다.
    if st.button("💾 구글 시트에 최종 발주 기록 저장", use_container_width=True):
        # 1. 추가발주수량을 입력했거나 2. 리오더 수량이 0이 아닌(수정된) 상품만 추출
        ready = df_display[
            (df_display['추가발주수량'] > 0) |   # 사장님이 직접 숫자를 넣은 발주 건
            (df_display['리오더 수량'] != 0)     # 리오더 수량에 변동이 있는 건
        ].copy()

        if not ready.empty:
            now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
            # 저장할 데이터 리스트 구성
            log_rows = [
                [
                    now_s, 
                    str(r[item]), 
                    str(r[option]), 
                    str(r[v_item]), 
                    int(r[avail]), 
                    int(r['리오더 수량']), 
                    int(r['추가발주수량']), 
                    int(r['권장 발주수량'])
                ] for _, r in ready.iterrows()
            ]
            
            try:
                # '발주기록' 시트에 데이터 추가
                get_sheet().worksheet("발주기록").append_rows(log_rows)
                st.success(f"✅ 추가 발주 및 변동 내역 {len(log_rows)}건 저장 완료!")
                time.sleep(1)
                st.rerun()
            except Exception as e:
                st.error(f"❌ 저장 실패: {e}")
        else:
            # 권장 수량만 있고 사장님이 아무것도 입력 안 했을 때 뜨는 메시지
            st.warning("⚠️ 기록할 '추가발주' 수량이나 '리오더' 변동 내역이 없습니다.")

with col_b2:
    # 다운로드는 실제 발주를 해야 하므로 (권장 + 추가) 합계가 있는 것들을 모두 포함합니다.
    df_display['합계'] = df_display['권장 발주수량'] + df_display['추가발주수량']
    csv_target = df_display[df_display['합계'] > 0]
    
    if not csv_target.empty:
        st.download_button(
            label="📥 최종 발주서 CSV 다운로드",
            data=csv_target[[item, option, v_item, avail, '리오더 수량', '추가발주수량', '권장 발주수량']].to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig'),
            file_name=f"발주서_{d5_d.strftime('%m%d')}.csv",
            use_container_width=True
        )


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





# ==========================================================
# --- [7단계: 실시간 전체 리오더 현황 (사장님 요청 컬럼 순서)] ---
# ==========================================================
if st.session_state.get('analyzed'):
    st.divider()
    st.subheader("📦 7단계: 실시간 전체 리오더 현황판")
    st.info("💡 현재 메인 시트(시트1)에서 '리오더 수량'이 남아있는 상품들을 정해진 순서대로 보여줍니다.")

    # [1. 메인 데이터 로드 및 정제]
    df_total = st.session_state.df_raw.copy()
    
    p = st.session_state.p
    it_col, op_col, vn_col, vi_col = p['it'], p['op'], p['vn'], p['vi']
    
    # 리오더 수량 숫자형 변환
    if "리오더 수량" in df_total.columns:
        df_total["리오더 수량"] = pd.to_numeric(df_total["리오더 수량"], errors='coerce').fillna(0).astype(int)
        
        # [2. 리오더 수량이 0보다 큰 상품만 필터링]
        df_reorder_active = df_total[df_total["리오더 수량"] > 0].copy()
        
        if not df_reorder_active.empty:
            # [3. 상단 요약 정보 계산]
            total_items = len(df_reorder_active)
            total_qty = df_reorder_active["리오더 수량"].sum()
            
            c1, c2 = st.columns(2)
            c1.metric("총 리오더 품목 수", f"{total_items}건")
            c2.metric("전체 리오더 총합계", f"{total_qty:,}개")
            
            # [4. 검색 UI]
            search_re = st.text_input("🔎 공급처 또는 상품명으로 검색", placeholder="검색어를 입력하세요...", key="search_re_active_v2")
            
            if search_re:
                df_reorder_active = df_reorder_active[
                    df_reorder_active[it_col].astype(str).str.contains(search_re, case=False) |
                    df_reorder_active[vn_col].astype(str).str.contains(search_re, case=False) |
                    df_reorder_active[vi_col].astype(str).str.contains(search_re, case=False)
                ]

            # [5. 컬럼 순서 재배치 및 이름 변경]
            # 사장님 요청 순서: 공급처 => 상품명 => 옵션 => 공급쳐상품명 => 현재리오더수량
            display_map = {
                vn_col: "공급처",
                it_col: "상품명",
                op_col: "옵션",
                vi_col: "공급처상품명",
                "리오더 수량": "현재 리오더 수량"
            }
            
            # 실제 존재하는 컬럼만 선별하여 순서대로 배치
            final_cols = [vn_col, it_col, op_col, vi_col, "리오더 수량"]
            df_reorder_display = df_reorder_active[final_cols].rename(columns=display_map)
            
            # [6. 데이터 에디터 출력 (공급처 -> 상품명 순으로 정렬)]
            st.data_editor(
                df_reorder_display.sort_values(by=["공급처", "상품명"]),
                use_container_width=True,
                hide_index=True,
                disabled=True, # 현황판이므로 수정 불가 모드
                column_config={
                    "공급처": st.column_config.TextColumn(width=100),
                    "상품명": st.column_config.TextColumn(width=300),
                    "옵션": st.column_config.TextColumn(width=120),
                    "공급처상품명": st.column_config.TextColumn(width=150),
                    "현재 리오더 수량": st.column_config.NumberColumn(width=100, format="%d개")
                }
            )
            
            # [7. 다운로드 버튼]
            st.download_button(
                label="📥 현재 리오더 현황 명단 다운로드 (CSV)",
                data=df_reorder_display.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig'),
                file_name=f"전체리오더현황_{datetime.now(KST).strftime('%m%d_%H%M')}.csv",
                mime="text/csv",
                use_container_width=True
            )
            
        else:
            st.success("✅ 현재 리오더 중인 상품이 없습니다.")
    else:
        st.error("'리오더 수량' 컬럼을 찾을 수 없습니다. 시트의 컬럼명을 확인해 주세요.")
