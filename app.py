import streamlit as st
import pandas as pd
import time
import io
import streamlit.components.v1 as components

# 🚨 [수정] pytz를 삭제하고 파이썬 기본 기능만 사용 (설치 에러 완벽 방지)
from datetime import datetime, timedelta, timezone

# 1. 한국 시간(KST) 및 오늘 날짜 설정 (pytz 없이도 정확함)
KST = timezone(timedelta(hours=9)) 
current_today = datetime.now(KST).date()

# 2. (선택사항) 만약 코드 하단에서 pytz를 꼭 써야 하는 상황이라면 
# 아래 try-except 구문을 써서 설치 안 됐을 때를 대비하세요.
try:
    import pytz
    KST_PYTZ = pytz.timezone('Asia/Seoul')
except ImportError:
    # pytz가 없으면 위에서 만든 기본 KST를 사용함
    pass


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


# [추가 함수] 구글 시트1에서 기존 '리오더 수량' 컬럼만 가져오기
def sync_reorder_from_sheet(df_uploaded):
    try:
        # 1. 시트1 열기
        sh = get_sheet()
        ws = sh.worksheet("시트1")
        all_data = ws.get_all_records()
        
        if all_data:
            df_sheet = pd.DataFrame(all_data)
            
            # 2. 필수 컬럼 확인 (상품명, 옵션, 리오더 수량)
            if "상품명" in df_sheet.columns and "옵션" in df_sheet.columns and "리오더 수량" in df_sheet.columns:
                # 시트에서 필요한 것만 추출하여 정리
                df_ref = df_sheet[['상품명', '옵션', '리오더 수량']].copy()
                df_ref['상품명'] = df_ref['상품명'].astype(str).str.strip()
                df_ref['옵션'] = df_ref['옵션'].astype(str).str.strip()
                df_ref['리오더 수량'] = pd.to_numeric(df_ref['리오더 수량'], errors='coerce').fillna(0).astype(int)
                
                # 3. 현재 업로드된 데이터(df_uploaded)와 병합
                # 엑셀의 기존 '리오더 수량'은 무시하고 시트의 최신값을 가져옵니다.
                if "리오더 수량" in df_uploaded.columns:
                    df_uploaded = df_uploaded.drop(columns=["리오더 수량"])
                
                df_final = pd.merge(df_uploaded, df_ref, on=['상품명', '옵션'], how='left')
                df_final['리오더 수량'] = df_final['리오더 수량'].fillna(0).astype(int)
                
                return df_final
        return df_uploaded
    except Exception as e:
        st.error(f"⚠️ 시트에서 리오더 수량을 동기화하는 중 오류: {e}")
        return df_uploaded


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



  # 데이터 로드 로직 (이 블록 전체를 교체하세요)
    if up_file:
        if st.session_state.get('df_raw') is None:
            try:
                # 1. 먼저 업로드한 파일을 읽습니다.
                df = pd.read_csv(up_file) if up_file.name.endswith('.csv') else pd.read_excel(up_file)
                df.columns = df.columns.str.strip()
                
                # 2. ⭐ [중요] 여기서 시트의 '리오더 수량'을 가져와 합칩니다!
                # 아까 정의한 함수를 여기서 호출(사용)해야 데이터가 0이 안 됩니다.
                with st.spinner("🔄 구글 시트에서 기존 리오더 수량을 동기화 중..."):
                    df = sync_reorder_from_sheet(df)
                
                # 3. '리오더 수량' 컬럼이 없는 경우를 대비한 기본 처리
                if "리오더 수량" not in df.columns: 
                    df["리오더 수량"] = 0
                
                df = df.fillna("") 
                st.session_state.df_raw = df
                st.success("✅ 파일 업로드 및 시트 데이터 동기화 완료!")
                
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



# ------------------------------------------------------------------
# [4단계: 데이터 편집 및 재고 관리] - 검색 로직 강화
# ------------------------------------------------------------------
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📊 4단계: 데이터 편집 및 재고 관리")

    # 매핑 변수 설정
    p = st.session_state.p
    sold_out_col, item, option = p['so'], p['it'], p['op']
    vendor, v_item = p['vn'], p['vi']
    stock, avail, t3day, t7day = p['st'], p['av'], p['t3'], p['t7']
    reg_date_col = p.get('reg') 
    lt, ss = p['lt'], p['ss']

    # 데이터 복사 및 숫자 변환
    df_work = st.session_state.df_raw.copy()
    for col in [stock, avail, t3day, t7day]:
        if col in df_work.columns:
            df_work[col] = pd.to_numeric(df_work[col], errors='coerce').fillna(0).astype(int)
    
    if "리오더 수량" not in df_work.columns: df_work["리오더 수량"] = 0
    df_work["리오더 수량"] = pd.to_numeric(df_work["리오더 수량"], errors='coerce').fillna(0).astype(int)
    if "비고" not in df_work.columns: df_work["비고"] = ""
    df_work["리오더 입고수량"] = 0 

    # UI 레이아웃
    f_c1, f_c2, f_c3 = st.columns([1, 2, 1])
    with f_c1: filter_m = st.selectbox("🚦 필터", ["전체보기", "정상만", "품절만"], index=1, key="v4_main_filter")
    with f_c2: search_q = st.text_input("🔍 검색", placeholder="검색할 상품을 넣어주세요...", key="v4_main_search")
    with f_c3: hist_date_4 = st.date_input("🗓️ 입고 기록 날짜", datetime.now(KST).date(), key="v4_main_date")

    # 판매량 및 권장발주 계산
    def calc_daily(row):
        t7, t3 = row[t7day], row[t3day]
        if reg_date_col and reg_date_col in row and pd.notnull(row[reg_date_col]):
            today = datetime.now(KST).date()
            reg_dt = row[reg_date_col].date() if hasattr(row[reg_date_col], 'date') else pd.to_datetime(row[reg_date_col]).date()
            days_diff = (today - reg_dt).days
            if 0 <= days_diff < 3:
                return int(round(t3 / (days_diff + 1))) if t3 > 0 else 0
        if t7 > 0: return int(round(t7 / 7))
        elif t3 > 0: return int(round(t3 / 3))
        return 0

    df_work['일판매'] = df_work.apply(calc_daily, axis=1)
    df_work['3일발주'] = (df_work['일판매'] * 3).astype(int)
    df_work['권장발주'] = ((df_work['일판매'] * (lt + ss)) - (df_work[avail] + df_work['리오더 수량'])).clip(lower=0).astype(int)

    # --- [검색 및 필터 로직 개선] ---
    if search_q:
        # 검색어가 있으면 필터와 상관없이 전체에서 검색
        df_f = df_work[df_work[item].astype(str).str.contains(search_q, case=False) | 
                       df_work[option].astype(str).str.contains(search_q, case=False)]
        if df_f.empty:
            st.warning(f"⚠️ '{search_q}' 상품을 찾을 수 없습니다. (전체 목록 표시)")
            # 결과 없으면 필터 조건대로 표시
            is_so = df_work[sold_out_col].astype(str).str.contains('품절', na=False)
            df_f = df_work[~is_so] if filter_m == "정상만" else (df_work[is_so] if filter_m == "품절만" else df_work)
    else:
        # 검색어 없으면 기존 필터 적용
        is_so = df_work[sold_out_col].astype(str).str.contains('품절', na=False)
        df_f = df_work[~is_so] if filter_m == "정상만" else (df_work[is_so] if filter_m == "품절만" else df_work)

    # 화면 출력용
    df_disp = df_f.rename(columns={sold_out_col: "상태", vendor: "공급쳐", v_item: "공급처상품명", item: "상품명", option: "옵션", stock: "정상", avail: "가용"})
    cols = ["상태", "공급쳐", "상품명", "옵션", "공급처상품명", "정상", "가용", "리오더 수량", "리오더 입고수량", "3일발주", "일판매", "권장발주"]

    with st.form("v4_master_form"):
        edited_v4 = st.data_editor(df_disp[cols], use_container_width=True, hide_index=True, key="v4_editor",
                                column_config={
                                    "상품명": st.column_config.TextColumn(width=300),
                                    "공급처상품명": st.column_config.TextColumn(width=150),
                                    "리오더 입고수량": st.column_config.NumberColumn("리오더 입고", format="%d", min_value=0)
                                })
        
        if st.form_submit_button("💾 데이터 저장 및 입고 반영", use_container_width=True, type="primary"):
            user_edits = st.session_state["v4_editor"].get("edited_rows", {})
            if user_edits:
                m_sh, h_sh = get_sheet().worksheet("시트1"), get_sheet().worksheet("입고기록")
                for r_idx, changes in user_edits.items():
                    idx = df_disp.index[int(r_idx)]
                    if "리오더 수량" in changes: 
                        st.session_state.df_raw.at[idx, "리오더 수량"] = int(changes["리오더 수량"])
                    if "리오더 입고수량" in changes:
                        qty = int(changes["리오더 입고수량"])
                        if qty > 0:
                            old_v = int(st.session_state.df_raw.at[idx, "리오더 수량"])
                            st.session_state.df_raw.at[idx, "리오더 수량"] = max(0, old_v - qty)
                            h_sh.append_row([datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S'), str(df_disp.at[idx, "상품명"]), str(df_disp.at[idx, "옵션"]), qty])
                
                df_save = st.session_state.df_raw.copy().fillna("").astype(str)
                m_sh.update([df_save.columns.values.tolist()] + df_save.values.tolist())
                st.success("✅ 메인 데이터 저장 완료!"); time.sleep(0.5); st.rerun()


# ------------------------------------------------------------------
# [5단계: 최종 발주 요약 및 구글 시트 저장] - 검색 강화 & 다운로드 포함
# ------------------------------------------------------------------
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📋 5단계: 최종 발주 리스트 요약")

    # 4단계 계산 데이터 활용 (품절 제외)
    df_v5 = df_work[~df_work[sold_out_col].astype(str).str.contains('품절', na=False)].copy()
    
    # 상단 필터 레이아웃
    f1, f2, f3 = st.columns([1.5, 2, 1])
    m5_f = f1.selectbox("🚦 상태 필터", ["🚨 고위험/주의", "✅ 전체정상"], key="v5_main_filter")
    s5_q = f2.text_input("🔍 상품명 검색", key="v5_main_search", placeholder="검색할 상품을 넣어주세요...")
    
    from datetime import datetime, timedelta, timezone
    KST_SAFE = timezone(timedelta(hours=9)) 
    current_today = datetime.now(KST_SAFE).date()
    d5_d = f3.date_input("🗓️ 기준 날짜", current_today, key="v5_main_date")

    if 'add_order_dict' not in st.session_state: 
        st.session_state.add_order_dict = {}
    
    # 기본 데이터 보정
    if '권장발주' not in df_v5.columns: df_v5['권장발주'] = 0
    df_v5['추가발주수량'] = df_v5.index.map(st.session_state.add_order_dict).fillna(0).astype(int)
    df_v5['상태'] = df_v5['권장발주'].apply(lambda x: "🚨 긴급" if x > 0 else "✅ 정상")

    # --- [검색 및 필터 로직 개선] ---
    # 검색어가 있으면 필터 무시하고 전체 발주 대상에서 검색
    if s5_q:
        df_v5_view = df_v5[df_v5[item].astype(str).str.contains(s5_q, case=False) | 
                           df_v5[option].astype(str).str.contains(s5_q, case=False)]
        
        if df_v5_view.empty:
            st.error(f"❌ '{s5_q}' 상품을 찾을 수 없습니다. (발주 대상 아님)")
            # 결과 없으면 필터 유지
            danger_names = df_v5[df_v5['권장발주'] > 0][item].unique() if item in df_v5.columns else []
            df_v5_view = df_v5[df_v5[item].isin(danger_names)].copy() if m5_f == "🚨 고위험/주의" else df_v5[~df_v5[item].isin(danger_names)].copy()
    else:
        # 검색어 없으면 필터링
        danger_names = df_v5[df_v5['권장발주'] > 0][item].unique() if item in df_v5.columns else []
        if m5_f == "🚨 고위험/주의":
            df_v5_view = df_v5[df_v5[item].isin(danger_names)].copy()
        else:
            df_v5_view = df_v5[~df_v5[item].isin(danger_names)].copy()
    
    df_v5_view = df_v5_view.sort_values(by=[item, option])

    # 에디터 매핑 및 출력
    full_map = {"상태": "상태", item: "상품명", option: "옵션", v_item: "공급처상품명", avail: "가용", "리오더 수량": "기존", "추가발주수량": "추가", "권장발주": "권장", "비고": "📝 이슈/입고메모"}
    available_map = {k: v for k, v in full_map.items() if k in df_v5_view.columns}
    
    with st.form("v5_master_form"):
        df_ed_v5 = df_v5_view[list(available_map.keys())].rename(columns=available_map)
        st.data_editor(df_ed_v5, use_container_width=True, hide_index=True, key="v5_editor",
            column_config={
                "상태": st.column_config.TextColumn("상태", width=70), 
                "상품명": st.column_config.TextColumn("상품명", width=250), 
                "옵션": st.column_config.TextColumn("옵션", width=150),
                "📝 이슈/입고메모": st.column_config.TextColumn(width=400)
            },
            disabled=[col for col in df_ed_v5.columns if col not in ["추가", "📝 이슈/입고메모"]]
        )
        
        if st.form_submit_button("✅ 1. 추가발주 및 메모 확정 (메인 반영)", use_container_width=True, type="primary"):
            edits_v5 = st.session_state["v5_editor"].get("edited_rows", {})
            if edits_v5:
                for r_idx, val in edits_v5.items():
                    idx = df_v5_view.index[int(r_idx)]
                    if "추가" in val:
                        st.session_state.df_raw.at[idx, "리오더 수량"] += int(val["추가"])
                        st.session_state.add_order_dict[idx] = int(val["추가"])
                    if "📝 이슈/입고메모" in val:
                        st.session_state.df_raw.at[idx, "비고"] = str(val["📝 이슈/입고메모"])
                
                m_sh = get_sheet().worksheet("시트1")
                df_save_m = st.session_state.df_raw.copy().fillna("").astype(str)
                m_sh.update([df_save_m.columns.values.tolist()] + df_save_m.values.tolist())
                st.success("✅ 메인 시트 저장 완료!"); time.sleep(0.5); st.rerun()
            else:
                st.info("💡 수정된 내용이 없습니다.")

    st.divider()
    c_save, c_down = st.columns(2)
    
    # [2. 구글 시트 전송]
    with c_save:
        if st.button("💾 2. 구글 시트에 최종 발주 기록 저장", use_container_width=True, type="primary"):
            valid_ids = [k for k, v in st.session_state.add_order_dict.items() if v > 0]
            if not valid_ids:
                st.warning("⚠️ 저장할 추가 수량이 없습니다. [1. 확정]을 먼저 눌러주세요.")
            else:
                try:
                    with st.spinner("🚀 발주기록 전송 중..."):
                        ws_log = get_sheet().worksheet("발주기록")
                        now_s = datetime.now(KST_SAFE).strftime('%Y-%m-%d %H:%M:%S')
                        log_rows = []
                        for idx in valid_ids:
                            if idx not in st.session_state.df_raw.index: continue
                            row = st.session_state.df_raw.loc[idx]
                            add_qty = st.session_state.add_order_dict[idx]
                            log_rows.append([
                                now_s, 
                                str(row.get(item, "")), 
                                str(row.get(option, "")), 
                                str(row.get(v_item, "")), 
                                int(row.get(avail, 0)), 
                                int(row.get('리오더 수량', 0) - int(add_qty)), 
                                int(add_qty), 
                                int(row.get('권장발주', 0)), 
                                str(row.get('비고', "")), 
                                str(row.get(vendor, ""))
                            ])
                        if log_rows:
                            ws_log.append_rows(log_rows)
                            st.session_state.add_order_dict = {}
                            st.success("✅ 시트 저장 성공!"); time.sleep(1); st.rerun()
                except Exception as e: 
                    st.error(f"❌ 저장 실패: {e}")

    # [3. CSV 다운로드] - 버튼 고정 노출 버전
    with c_down:
        # 권장발주 + 추가발주 합산 계산
        df_v5['최종합계'] = df_v5.get('권장발주', 0) + df_v5['추가발주수량']
        csv_target = df_v5[df_v5['최종합계'] > 0].copy()
        
        # 출력할 컬럼 설정 (데이터가 없어도 헤더는 나오게 준비)
        down_cols = [vendor, item, option, v_item, '최종합계']
        existing_cols = [c for c in down_cols if c in csv_target.columns]
        
        if not csv_target.empty:
            # 1. 발주 수량이 있을 때: 정상 데이터 생성
            csv_res = csv_target[existing_cols].rename(columns={
                vendor: "공급처", item: "상품명", option: "옵션", v_item: "공급처상품명", "최종합계": "발주수량"
            })
            button_label = "📥 최종 발주서 CSV 다운로드"
            is_disabled = False
        else:
            # 2. 발주 수량이 없을 때: 헤더만 있는 빈 데이터 생성 (버전 유지용)
            # 빈 데이터프레임 생성
            csv_res = pd.DataFrame(columns=["공급처", "상품명", "옵션", "공급처상품명", "발주수량"])
            button_label = "📥 다운로드 (발주 수량 없음)"
            is_disabled = True # 버튼을 비활성화(클릭 안됨) 상태로 노출하거나, False로 두면 빈 파일이 받아집니다.

        # 한글 깨짐 방지 인코딩
        csv_data = csv_res.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
        
        # 버튼 고정 노출 (disabled 옵션으로 제어)
        st.download_button(
            label=button_label,
            data=csv_data,
            file_name=f"발주서_{d5_d.strftime('%m%d')}_빈파일.csv" if csv_target.empty else f"발주서_{d5_d.strftime('%m%d')}.csv",
            mime="text/csv",
            use_container_width=True,
            disabled=is_disabled  # 수량이 없으면 버튼은 보이지만 클릭은 안 되게 설정
        )



# ------------------------------------------------------------------
# [6단계: 전체 히스토리 관리] - 회차별 선택 + '일자별 전체 합계' 추가
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.subheader("📜 6단계: 전체 히스토리 관리")

    # [1] 상단 컨트롤러 (항상 노출)
    f1, f2, f3, f4 = st.columns([1.2, 0.8, 1.5, 1.5])
    
    with f1:
        today = datetime.now(KST).date()
        d_range = st.date_input("🗓️ 1. 조회 범위", value=(today, today), key="v6_date_range")
    
    with f2:
        st.write(""); st.write("") 
        search_trigger = st.button("🔍 2. 내역 조회", use_container_width=True, type="primary")

    if 'v6_data' not in st.session_state: st.session_state.v6_data = None
    if 'v6_sessions' not in st.session_state: st.session_state.v6_sessions = []

    # [데이터 로드]
    if search_trigger:
        try:
            with st.spinner("📡 데이터를 불러오는 중..."):
                worksheet = get_sheet().worksheet("발주기록")
                all_h = worksheet.get_all_values()
                if len(all_h) > 1:
                    df_all = pd.DataFrame(all_h[1:])
                    # 매핑 고정: 9번 업체명, 3번 공급처상품명
                    df_all.columns = ["발주시간", "상품명", "옵션", "공급처상품명", "가용", "기존", "추가", "권장", "이슈/메모", "업체명"]
                    
                    df_all["날짜_만"] = df_all["발주시간"].astype(str).str.slice(0, 10)
                    if isinstance(d_range, tuple) and len(d_range) == 2:
                        s_d, e_d = d_range[0].strftime('%Y-%m-%d'), d_range[1].strftime('%Y-%m-%d')
                        df_all = df_all[(df_all["날짜_만"] >= s_d) & (df_all["날짜_만"] <= e_d)]
                    
                    st.session_state.v6_data = df_all
                    st.session_state.v6_sessions = sorted(df_all["발주시간"].unique(), reverse=True)
                else:
                    st.session_state.v6_data = None
                    st.session_state.v6_sessions = []
                    st.info("💡 저장된 내역이 없습니다.")
        except Exception as e:
            st.error(f"📡 오류: {e}")

    with f3:
        h_q = st.text_input("🔍 3. 상품명 검색", placeholder="상품명 입력...", key="v6_search_q")
    
    with f4:
        if st.session_state.v6_sessions:
            # 🚨 "일자별 전체 합계" 옵션을 가장 위에 추가
            session_options = ["📊 선택 범위 전체 합계"] + [f"{len(st.session_state.v6_sessions)-i}회차 ({t[11:16]} 저장)" for i, t in enumerate(st.session_state.v6_sessions)]
            sel_session_label = st.selectbox("📦 4. 회차 선택", session_options, key="v6_session_select")
        else:
            st.selectbox("📦 4. 회차 선택", ["조회 결과 없음"], disabled=True, key="v6_session_select_empty")
            sel_session_label = None

    # [2] 결과 출력 영역
    if st.session_state.v6_data is not None and sel_session_label:
        df_raw_hist = st.session_state.v6_data.copy()
        
        # 필터링 로직
        if sel_session_label == "📊 선택 범위 전체 합계":
            # 날짜 범위 내 모든 데이터 합산 (상품명, 옵션, 업체명 기준)
            df_display = df_raw_hist.groupby(["업체명", "상품명", "옵션", "공급처상품명"], as_index=False).agg({
                "발주시간": "max",
                "가용": "last",
                "기존": "last",
                "추가": lambda x: pd.to_numeric(x, errors='coerce').sum(),
                "권장": "last",
                "이슈/메모": lambda x: " / ".join(set(filter(None, x)))
            })
            display_title = f"🗓️ {d_range[0]} ~ {d_range[1]} 전체 발주 합계"
        else:
            # 특정 회차만 필터링
            target_time = st.session_state.v6_sessions[session_options.index(sel_session_label)-1]
            df_display = df_raw_hist[df_raw_hist["발주시간"] == target_time].copy()
            display_title = f"✅ {sel_session_label} 상세 내역"

        # 상품명 검색 필터 적용
        if h_q:
            df_display = df_display[df_display["상품명"].astype(str).str.contains(h_q, case=False)]

        if not df_display.empty:
            st.write(f"#### {display_title}")
            
            # 표시 순서 고정
            display_order = ["발주시간", "업체명", "상품명", "옵션", "공급처상품명", "가용", "기존", "추가", "권장", "이슈/메모"]
            
            st.dataframe(
                df_display[display_order],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "발주시간": st.column_config.TextColumn("📅 마지막저장", width=160),
                    "업체명": st.column_config.TextColumn("🏭 업체명", width=120),
                    "상품명": st.column_config.TextColumn("📦 상품명", width=250),
                    "옵션": st.column_config.TextColumn("옵션", width=150),
                    "공급처상품명": st.column_config.TextColumn("🆔 공급처 상품명", width=180),
                    "가용": st.column_config.TextColumn("가용", width=80),
                    "기존": st.column_config.TextColumn("기존", width=80),
                    "추가": st.column_config.TextColumn("추가", width=80),
                    "권장": st.column_config.TextColumn("권장", width=80),
                    "이슈/메모": st.column_config.TextColumn("📝 이슈/메모", width=450)
                }
            )
            
             # 다운로드 버튼
            csv_data = df_display[display_order].to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button(f"📥 {sel_session_label} CSV 다운로드", csv_data, f"발주합계_{datetime.now().strftime('%m%d')}.csv", use_container_width=True)




# ------------------------------------------------------------------
# [7단계: 실시간 리오더 누적 상황판] - 순서: 기간 > 갱신 > 검색 > 업체선택
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.subheader("🚀 7단계: 실시간 리오더 누적 및 상황판")

    try:
        # [데이터 실시간 로드]
        ws_v7 = get_sheet().worksheet("발주기록")
        all_v7 = ws_v7.get_all_values()
        
        if len(all_v7) > 1:
            df_raw_v7 = pd.DataFrame(all_v7[1:])
            # 🚨 위치 고정: 0:시간, 1:상품명, 2:옵션, 3:거래처상품명, 6:추가수량, 9:업체명
            df_raw_v7.columns = ["발주시간", "상품명", "옵션", "거래처상품명", "가용", "기존", "추가", "권장", "이슈/메모", "업체명"]
            
            # 숫자 변환 및 날짜 추출
            df_raw_v7["추가"] = pd.to_numeric(df_raw_v7["추가"], errors='coerce').fillna(0)
            df_raw_v7["날짜"] = df_raw_v7["발주시간"].astype(str).str.slice(0, 10)
            
            # [상단 상황판 보드] - 오늘(Today) 기준 실시간 요약
            today_str = datetime.now(KST).strftime('%Y-%m-%d')
            df_today = df_raw_v7[df_raw_v7["날짜"] == today_str]

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("📅 기준 날짜", today_str)
            c2.metric("📦 총 발주 상품수", f"{df_today['상품명'].nunique()}종")
            c3.metric("🏭 발주 업체수", f"{df_today['업체명'].nunique()}곳")
            c4.metric("🔢 총 누적수량", f"{int(df_today['추가'].sum())}개")
            st.divider()

            # [필터 레이아웃] 사장님 요청 순서: 기간 -> 갱신 -> 검색 -> 업체선택
            f1, f2, f3, f4 = st.columns([1.2, 0.8, 1.5, 1.5])
            
            with f1:
                # 1. 기간 선택
                d_range_v7 = st.date_input("🗓️ 1. 기간 선택", value=(today, today), key="v7_range")
            
            with f2:
                st.write(""); st.write("") # 높이 맞춤
                # 2. 데이터 갱신 버튼
                btn_v7 = st.button("📈 2. 데이터 갱신", use_container_width=True, type="primary")

            with f3:
                # 3. 상품명 검색
                q_v7 = st.text_input("🔍 3. 상품명 검색", placeholder="상품명 입력...", key="v7_search_q")

            with f4:
                # 4. 업체 선택 셀렉트박스
                v_list = sorted(df_raw_v7["업체명"].unique().tolist())
                v_choice = st.selectbox("🏭 4. 업체 선택", ["전체 업체"] + v_list, key="v7_vendor_sel")

            # [필터링 로직 적용]
            # 날짜 범위 필터
            if isinstance(d_range_v7, tuple) and len(d_range_v7) == 2:
                s_d, e_d = d_range_v7[0].strftime('%Y-%m-%d'), d_range_v7[1].strftime('%Y-%m-%d')
                df_filtered = df_raw_v7[(df_raw_v7["날짜"] >= s_d) & (df_raw_v7["날짜"] <= e_d)]
            else:
                df_filtered = df_today

            # 상품명 검색 필터
            if q_v7:
                df_filtered = df_filtered[df_filtered["상품명"].astype(str).str.contains(q_v7, case=False)]

            # 업체 선택 필터
            if v_choice != "전체 업체":
                df_filtered = df_filtered[df_filtered["업체명"] == v_choice]

            # [데이터 그룹화 집계]
            df_final = df_filtered.groupby(["날짜", "업체명", "상품명", "옵션", "거래처상품명"], as_index=False).agg({
                "추가": "sum", # 리오더 수량 누적 합산
                "이슈/메모": lambda x: " / ".join(set(filter(None, x))) # 메모 통합
            })

            if not df_final.empty:
                # 🚨 출력 순서 고정: 날짜 > 업체명 > 상품명 > 옵션 > 거래처상품명
                display_order = ["날짜", "업체명", "상품명", "옵션", "거래처상품명", "추가", "이슈/메모"]
                
                st.dataframe(
                    df_final[display_order],
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "날짜": st.column_config.TextColumn("📅 날짜", width=110),
                        "업체명": st.column_config.TextColumn("🏭 업체명", width=120),
                        "상품명": st.column_config.TextColumn("📦 상품명", width=250),
                        "옵션": st.column_config.TextColumn("옵션", width=150),
                        "거래처상품명": st.column_config.TextColumn("🆔 거래처상품명", width=180),
                        "추가": st.column_config.NumberColumn("🔢 누적수량", width=80),
                        "이슈/메모": st.column_config.TextColumn("📝 통합메모", width=400)
                    }
                )
                
                # CSV 다운로드
                csv_v7 = df_final[display_order].to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
                st.download_button("📥 실시간 누적 집계표(CSV) 다운로드", csv_v7, f"최종집계_{today_str}.csv", use_container_width=True)
            else:
                st.warning("🧐 해당 조건에 맞는 데이터가 없습니다.")

        else:
            st.info("💡 저장된 발주 데이터가 없습니다.")
            
    except Exception as e:
        st.error(f"📡 데이터 로딩 오류: {e}")




# --- [🌙 탭 2: 동대문 사입 관리] ---

with tab2:

    st.subheader("🌙 동대문 사입 및 미납 관리")

    dong_file = st.file_uploader("동대문 주문 리스트 업로드", type=['xlsx', 'csv'], key="dong_tab_upload")

    if dong_file:

        if "last_file_name" not in st.session_state or st.session_state.last_file_name != dong_file.name:

            df = pd.read_excel(dong_file)

            df.columns = df.columns.str.strip()

            required_cols = ['선택', '품절', '상품명', '공급처', '공급처상품명', '정상재고', '가용재고', '판매수량', '발주수량', '가중율', '3일판매']

            for col in required_cols:

                if col not in df.columns: df[col] = 0 if col not in ['선택', '품절', '상품명', '공급처', '공급처상품명'] else ""

            for col in ['정상재고', '가용재고', '3일판매']: df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            df['판매수량'] = (df['정상재고'] - df['가용재고']).clip(lower=0)

            df['가중율'] = df['판매수량'].apply(lambda n: 2.0 if n >= 10 else (1.5 if n >= 6 else (1.2 if n >= 3 else 1.0)))

            df['발주수량'] = (df['판매수량'] * df['가중율']).astype(int)

            st.session_state.df_dong_current = df[required_cols]

            st.session_state.last_file_name = dong_file.name



        df_display = st.session_state.df_dong_current.copy()

        search_query = st.text_input("상품명 검색 (사입)")

        if search_query: df_display = df_display[df_display['상품명'].astype(str).str.contains(search_query, case=False, na=False)]

        

        df_display['선택'] = df_display['선택'].astype(bool)

        edited_df = st.data_editor(df_display, use_container_width=True, key="dong_editor")

        

        st.divider()

        c1, c2, c3 = st.columns(3)

        add_val = c1.number_input("추가 수량", value=1, min_value=1)

        if c2.button("🚀 선택 상품 수량 더하기"):

            selected = edited_df[edited_df['선택'] == True].index

            for idx in selected: st.session_state.df_dong_current.at[idx, '발주수량'] += add_val

            st.rerun()

        csv = edited_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')

        c3.download_button("📥 엑셀 다운로드", csv, "사입리스트.csv", "text/csv")
