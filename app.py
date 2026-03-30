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

# [공통 설정] - 코드 최상단 session_state 선언부
if 'common_search' not in st.session_state:
    st.session_state.common_search = ""
if 'add_order_dict' not in st.session_state:
    st.session_state.add_order_dict = {}

# [통합 함수] 검색어 동기화 + 5단계 수량 실시간 보존
def sync_all_and_save_mem():
    """
    검색을 하거나 화면이 바뀔 때 실행됩니다.
    에디터에 입력된 '추가입력' 수치를 금고(add_order_dict)에 즉시 박아넣습니다.
    """
    # 1. 5단계 추가수량 실시간 보존
    if "v5_editor_fixed" in st.session_state:
        edits = st.session_state["v5_editor_fixed"].get("edited_rows", {})
        if "current_v5_index" in st.session_state and edits:
            v_idx_list = st.session_state.current_v5_index
            for r_idx_str, val in edits.items():
                actual_idx = v_idx_list[int(r_idx_str)]
                # '추가입력' 칸에 숫자가 있으면 금고에 저장 (키값은 5단계 d_map과 일치해야 함)
                if "추가입력" in val:
                    st.session_state.add_order_dict[actual_idx] = int(val["추가입력"])

    # 2. 검색어 동기화
    if "v4_fix_search" in st.session_state:
        st.session_state.common_search = st.session_state.v4_fix_search
    if "v5_search_fixed" in st.session_state:
        st.session_state.common_search = st.session_state.v5_search_fixed

# 3. 일판매량 계산 함수 (신상품 등록일 보정 포함)
def calc_daily_sales_with_reg_v4(row, p_config):
    t7day, t3day = p_config['t7'], p_config['t3']
    reg_date_col = p_config.get('reg')
    t7, t3 = row[t7day], row[t3day]
    
    if reg_date_col and reg_date_col in row and pd.notnull(row[reg_date_col]):
        today = datetime.now(KST).date()
        reg_dt = row[reg_date_col].date() if hasattr(row[reg_date_col], 'date') else pd.to_datetime(row[reg_date_col]).date()
        days_diff = (today - reg_dt).days
        if 0 <= days_diff < 3:
            actual_days = days_diff + 1
            return int(round(t3 / actual_days)) if t3 > 0 else 0
            
    if t7 > 0: return int(round(t7 / 7))
    elif t3 > 0: return int(round(t3 / 3))
    return 0

def sync_all_and_save_mem():
    """
    화면이 바뀌기 직전(검색, 필터 변경 등)에 
    현재 에디터에 떠 있는 모든 '추가입력' 수치를 금고(session_state)에 저장합니다.
    """
    # 5단계 에디터 데이터 처리
    if "v5_editor_fixed" in st.session_state:
        edits = st.session_state["v5_editor_fixed"].get("edited_rows", {})
        if "current_v5_index" in st.session_state and edits:
            v_idx_list = st.session_state.current_v5_index
            for r_idx_str, val in edits.items():
                try:
                    actual_idx = v_idx_list[int(r_idx_str)]
                    # 사용자가 입력한 컬럼명이 '추가입력' 또는 '발주추가' 등 무엇이든 대응
                    # 에디터 내부의 수정된 값을 추출
                    new_qty = None
                    if "추가입력" in val: new_qty = val["추가입력"]
                    elif "발주추가" in val: new_qty = val["발주추가"]
                    
                    if new_qty is not None:
                        st.session_state.add_order_dict[actual_idx] = int(new_qty)
                except Exception:
                    continue

    # 검색어 동기화
    if "v5_search_fixed" in st.session_state:
        st.session_state.common_search = st.session_state.v5_search_fixed


def sync_all_and_save_mem():
    """
    검색어를 입력하거나 필터를 바꿀 때 실행되어, 
    현재 화면(에디터)에 있는 '발주추가' 수량을 금고(add_order_dict)에 즉시 저장합니다.
    """
    # 5단계 에디터의 수정사항 감지
    if "v5_editor_fixed" in st.session_state:
        edits = st.session_state["v5_editor_fixed"].get("edited_rows", {})
        # 현재 화면에 보이는 행들이 원본의 몇 번째 인덱스인지 확인
        if "current_v5_index" in st.session_state and edits:
            v_idx_list = st.session_state.current_v5_index
            for r_idx_str, val in edits.items():
                actual_idx = v_idx_list[int(r_idx_str)]
                # '발주추가' 컬럼에 숫자가 입력되었다면 금고에 기록
                if "발주추가" in val:
                    st.session_state.add_order_dict[actual_idx] = int(val["발주추가"])

    # 검색어 상태 동기화
    if "v5_search_fixed" in st.session_state:
        st.session_state.common_search = st.session_state.v5_search_fixed

def add_to_order():
    # 개별 입력창에서 넣은 값을 금고에 즉시 저장
    if "temp_qty" in st.session_state and "selected_item_idx" in st.session_state:
        idx = st.session_state.selected_item_idx
        qty = st.session_state.temp_qty
        st.session_state.add_order_dict[idx] = qty
        st.toast(f"✅ 수량 {qty}개 임시 저장됨!")


def sync_all_and_save_mem():
    """
    검색을 하거나 필터를 바꿀 때 실행됩니다.
    에디터의 '수정 중인 데이터'를 강제로 긁어와서 금고(session_state)에 저장합니다.
    """
    # 5단계 에디터 데이터 처리
    if "v5_editor_fixed" in st.session_state:
        edits = st.session_state["v5_editor_fixed"].get("edited_rows", {})
        if "current_v5_index" in st.session_state and edits:
            v_idx_list = st.session_state.current_v5_index
            for r_idx_str, val in edits.items():
                try:
                    actual_idx = v_idx_list[int(r_idx_str)]
                    # '추가입력' 컬럼의 수정된 값을 추출하여 금고에 저장
                    if "추가입력" in val:
                        st.session_state.add_order_dict[actual_idx] = int(val["추가입력"])
                except:
                    continue

    # 검색어 동기화
    if "v5_search_fixed" in st.session_state:
        st.session_state.common_search = st.session_state.v5_search_fixed





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
# --- [4단계: 데이터 편집 및 재고 관리 (에러 방지 강화)] ---
# ==========================================================
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📊 4단계: 데이터 편집 및 재고 관리")

    # 1. 설정값 불러오기
    p = st.session_state.p
    sold_out_col, item, option = p['so'], p['it'], p['op']
    vendor, v_item = p['vn'], p['vi']
    stock, avail, t3day, t7day = p['st'], p['av'], p['t3'], p['t7']
    reg_date_col = p.get('reg')
    lt, ss = p['lt'], p['ss']

    # 2. 데이터 준비 및 숫자 타입 변환
    df_work = st.session_state.df_raw.copy()
    num_cols = [stock, avail, t3day, t7day]
    for col in num_cols:
        if col in df_work.columns:
            df_work[col] = pd.to_numeric(df_work[col], errors='coerce').fillna(0).astype(int)
    
    if "리오더 수량" not in df_work.columns: 
        df_work["리오더 수량"] = 0
    df_work["리오더 수량"] = pd.to_numeric(df_work["리오더 수량"], errors='coerce').fillna(0).astype(int)
    df_work["리오더 입고수량"] = 0 

    # 3. ⭐ 입고 이력 합산 (KeyError 방지 로직 포함)
    @st.cache_data(ttl=60)
    def get_in_history_v4_safe():
        try:
            sh_h = get_sheet().worksheet("입고기록")
            h_data = sh_h.get_all_records()
            if h_data:
                h_df = pd.DataFrame(h_data)
                h_df.columns = h_df.columns.str.strip() # 컬럼명 공백 제거
                # 필수 컬럼 존재 여부 확인
                if '상품명' in h_df.columns and '옵션' in h_df.columns and '입고수량' in h_df.columns:
                    return h_df.groupby(['상품명', '옵션'])['입고수량'].sum().reset_index()
            return pd.DataFrame(columns=['상품명', '옵션', '입고수량'])
        except Exception: 
            return pd.DataFrame(columns=['상품명', '옵션', '입고수량'])

    in_sum = get_in_history_v4_safe()
    
    # 안전한 병합(Merge) 수행
    if not in_sum.empty:
        df_work = pd.merge(df_work, in_sum.rename(columns={"입고수량":"과거리오더 입고"}), 
                           left_on=[item, option], right_on=['상품명', '옵션'], how="left").fillna(0)
    else:
        df_work["과거리오더 입고"] = 0

    # 4. 일판매량 및 권장발주 계산 (공통 함수 사용)
    # 상단에 calc_daily_sales_with_reg_v4 함수가 정의되어 있어야 합니다.
    df_work['일판매'] = df_work.apply(lambda x: calc_daily_sales_with_reg_v4(x, p), axis=1)
    df_work['3일발주'] = (df_work['일판매'] * 3).astype(int)
    df_work['권장발주'] = ((df_work['일판매'] * (lt + ss)) - (df_work[avail] + df_work['리오더 수량'])).clip(lower=0).astype(int)

    # 5. 상단 레이아웃 (필터 및 검색)
    f_c1, f_c2, f_c3 = st.columns([1, 2, 1])
    with f_c1: 
        filter_m = st.selectbox("🚦 필터", ["전체보기", "정상만", "품절만"], index=1, key="v4_fix_filter")
    with f_c2: 
        st.text_input("🔍 통합 검색 (상품명/옵션)", 
                     value=st.session_state.get('common_search', ""), 
                     key="v4_fix_search", 
                     on_change=sync_all_and_save_mem) # 수량 보존 및 검색 동기화
    with f_c3: 
        hist_date_4 = st.date_input("🗓️ 입고 날짜", datetime.now(KST).date(), key="v4_fix_date")

    # 6. 검색 및 필터링 실제 적용
    is_soldout = df_work[sold_out_col].astype(str).str.contains('품절', na=False)
    df_filtered = df_work[~is_soldout] if filter_m == "정상만" else (df_work[is_soldout] if filter_m == "품절만" else df_work)
    
    q = st.session_state.get('common_search', "")
    if q:
        df_filtered = df_filtered[
            df_filtered[item].astype(str).str.contains(q, case=False, na=False) | 
            df_filtered[option].astype(str).str.contains(q, case=False, na=False)
        ]

    # 7. 화면 출력용 데이터 가공
    df_display = df_filtered.rename(columns={
        sold_out_col: "상태", vendor: "공급쳐", v_item: "공급상품명", 
        item: "상품명", option: "옵션", stock: "정상", avail: "가용"
    })
    
    final_cols = ["상태", "공급쳐", "상품명", "옵션", "공급상품명", "정상", "가용", "리오더 수량", "리오더 입고수량", "과거리오더 입고", "3일발주", "일판매", "권장발주"]

    # 8. 데이터 에디터 및 저장 폼
    with st.form("v4_fix_master_form"):
        # 에러 방지를 위해 실제 존재하는 컬럼만 선택
        v_cols = [c for c in final_cols if c in df_display.columns]
        
        edited_v4 = st.data_editor(
            df_display[v_cols], 
            use_container_width=True, 
            hide_index=True, 
            key="v4_editor_fix",
            column_config={
                "상태": st.column_config.TextColumn(width=60),
                "상품명": st.column_config.TextColumn(width=320),
                "리오더 입고수량": st.column_config.NumberColumn("입고입력", width=70, format="%d", min_value=0),
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
                            old_val = int(st.session_state.df_raw.at[target_idx, "리오더 수량"])
                            # 리오더 수량에서 차감
                            st.session_state.df_raw.at[target_idx, "리오더 수량"] = max(0, old_val - in_qty)
                            # 입고 기록 시트 추가
                            h_sh.append_row([save_time, str(df_display.at[target_idx, "상품명"]), str(df_display.at[target_idx, "옵션"]), in_qty])

                # 구글 시트 반영
                df_to_save = st.session_state.df_raw.copy().fillna("").astype(str)
                m_sh.update([df_to_save.columns.values.tolist()] + df_to_save.values.tolist())
                st.success("✅ 저장 및 입고 기록이 완료되었습니다!"); time.sleep(0.5); st.rerun()



# ==========================================================
# --- [5단계: 최종 발주 리스트 (거래처/디자인/기록/다운로드 완결)] ---
# ==========================================================
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📋 5단계: 최종 발주 리스트 (거래처별 마스터 관리)")

    # [1. 데이터 기초 공사 및 지표 계산]
    p = st.session_state.p
    sold_out_col, item, option, v_item = p['so'], p['it'], p['op'], p['vi']
    stock, avail, t3day, t7day = p['st'], p['av'], p['t3'], p['t7']
    lt, ss = p['lt'], p['ss']

    df_v5 = st.session_state.df_raw.copy()
    
    # 숫자형 변환 (계산 오류 방지)
    for c in [stock, avail, t3day, t7day, '리오더 수량']:
        if c in df_v5.columns:
            df_v5[c] = pd.to_numeric(df_v5[c], errors='coerce').fillna(0).astype(int)
    
    if "리오더 수량" not in df_v5.columns: df_v5["리오더 수량"] = 0

    # [2. 입고 히스토리 실시간 합산]
    @st.cache_data(ttl=60)
    def get_v5_history_full():
        try:
            sh_h = get_sheet().worksheet("입고기록")
            h_data = sh_h.get_all_records()
            if h_data:
                h_df = pd.DataFrame(h_data)
                return h_df.groupby(['상품명', '옵션'])['입고수량'].sum().reset_index()
            return pd.DataFrame(columns=['상품명', '옵션', '입고수량'])
        except: return pd.DataFrame(columns=['상품명', '옵션', '입고수량'])

    df_h_v5 = get_v5_history_full().rename(columns={"입고수량": "최근입고합계"})
    df_v5 = pd.merge(df_v5, df_h_v5, left_on=[item, option], right_on=['상품명', '옵션'], how="left").fillna(0)
    df_v5.index = st.session_state.df_raw.index

    # [3. 핵심 계산 (신상품 보정 판매량 & 상태 아이콘)]
    df_v5['일판매량'] = df_v5.apply(lambda x: calc_daily_sales_with_reg_v4(x, p), axis=1)
    df_v5['3일분'] = (df_v5['일판매량'] * 3).astype(int)
    df_v5['권장수량'] = ((df_v5['일판매량'] * (lt + ss)) - (df_v5[avail] + df_v5['리오더 수량'])).clip(lower=0).astype(int)
    
    # 금고(Session State) 입력값 매핑 (데이터 증발 방지)
    df_v5['추가입력'] = df_v5.index.map(st.session_state.get('add_order_dict', {})).fillna(0).astype(int)

    def label_v5_status(row):
        if row['권장수량'] >= 10: return "🚨 고위험"
        elif row['권장수량'] > 0: return "⚠️ 주의"
        elif row['추가입력'] > 0: return "📝 입력중"
        return "✅ 정상"
    df_v5['상태분류'] = df_v5.apply(label_v5_status, axis=1)

    # [4. 디자인 및 필터 섹션 (친구/상태 필터)]
    with st.container(border=True):
        c_f1, c_f2, c_f3 = st.columns([1.5, 1.5, 1])
        with c_f1:
            vendor_list = ["전체 거래처"] + sorted(df_v5[v_item].unique().tolist())
            sel_vendor = st.selectbox("🤝 거래처(친구) 필터", vendor_list)
        with c_f2:
            st_filter = st.selectbox("🚦 상태 필터", ["🔥 발주필요(고위험/주의)", "📝 수동입력 중", "✅ 전체보기"])
        with c_f3:
            v5_date = st.date_input("🗓️ 발주 기준일", datetime.now(KST).date())

    # [5. 🛠️ 절대 안 날아가는 검색 & 입력 도구]
    st.markdown("#### 🔍 상품 검색 및 수량 확정")
    with st.container(border=True):
        c_search, c_val, c_save = st.columns([2, 1, 1])
        with c_search:
            q_secure = st.text_input("🔎 검색 (거래처 필터 영향 받음)", key="v5_q_secure").strip()
        
        # 필터링 로직
        search_df = df_v5.copy()
        if sel_vendor != "전체 거래처":
            search_df = search_df[search_df[v_item] == sel_vendor]
            
        if q_secure:
            res = search_df[search_df[item].str.contains(q_secure, case=False, na=False) | 
                            search_df[option].str.contains(q_secure, case=False, na=False)]
            if not res.empty:
                opts = {idx: f"[{r[v_item]}] {r[item]} | {r[option]} (권장:{r['권장수량']})" for idx, r in res.iterrows()}
                sel_idx = st.selectbox("품목 확정", opts.keys(), format_func=lambda x: opts[x])
                with c_val:
                    cur_v = int(st.session_state.add_order_dict.get(sel_idx, 0))
                    new_v = st.number_input("추가수량", min_value=0, value=cur_v, key="v5_input_val")
                with c_save:
                    st.write(" ")
                    if st.button("📥 저장", use_container_width=True, type="secondary"):
                        st.session_state.add_order_dict[sel_idx] = new_v
                        st.toast(f"✅ {new_v}개 저장됨!"); time.sleep(0.3); st.rerun()
            else:
                st.warning("검색 결과가 없습니다.")

    st.divider()

    # [6. 메인 작업 테이블 (디자인 강화)]
    display_df = df_v5.copy()
    if sel_vendor != "전체 거래처":
        display_df = display_df[display_df[v_item] == sel_vendor]
    
    if st_filter == "🔥 발주필요(고위험/주의)":
        display_df = display_df[display_df['권장수량'] > 0]
    elif st_filter == "📝 수동입력 중":
        display_df = display_df[display_df['추가입력'] > 0]

    d_map = {
        "상태분류": "상태", v_item: "거래처", item: "상품명", option: "옵션", 
        avail: "가용", "리오더 수량": "리중", "추가입력": "추가", "권장수량": "권장", 
        "3일분": "3일분", "최근입고합계": "최근입고"
    }
    
    st.dataframe(
        display_df[list(d_map.keys())].rename(columns=d_map),
        use_container_width=True, 
        hide_index=True,
        column_config={
            "상태": st.column_config.TextColumn(width=85),
            "거래처": st.column_config.TextColumn(width=100),
            "상품명": st.column_config.TextColumn(width=250),
        }
    )

    # [7. 하단 기능 버튼 (저장/히스토리/다운로드)]
    st.write("---")
    btn_c1, btn_c2, btn_c3 = st.columns(3)
    
    with btn_c1:
        # 1. 구글 시트 저장 (시트1 리오더 수량 합산)
        if st.button("🚀 시트1 최종 반영", use_container_width=True, type="primary"):
            if st.session_state.add_order_dict:
                m_sh = get_sheet().worksheet("시트1")
                for idx, qty in st.session_state.add_order_dict.items():
                    st.session_state.df_raw.at[idx, "리오더 수량"] += qty
                save_df = st.session_state.df_raw.copy().fillna("").astype(str)
                m_sh.update([save_df.columns.values.tolist()] + save_df.values.tolist())
                st.session_state.add_order_dict = {} # 초기화
                st.success("✅ 시트1 반영 완료!"); time.sleep(1); st.rerun()
            else:
                st.warning("입력된 수량이 없습니다.")

    with btn_c2:
        # 2. 발주 기록지(히스토리) 저장
        if st.button("💾 발주 히스토리 저장", use_container_width=True):
            log_df = df_v5[df_v5['추가입력'] > 0].copy()
            if not log_df.empty:
                now_str = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
                log_data = [[now_str, str(r[item]), str(r[option]), str(r[v_item]), int(r[avail]), int(r['리오더 수량']), int(r['추가입력']), int(r['권장수량'])] for _, r in log_df.iterrows()]
                get_sheet().worksheet("발주기록").append_rows(log_data)
                st.success("✅ 발주기록 저장 완료!")
            else:
                st.warning("입력 중인 상품이 없습니다.")

    with btn_c3:
        # 3. 엑셀(CSV) 다운로드
        df_v5['최종발주합계'] = df_v5['권장수량'] + df_v5['추가입력']
        csv_data = df_v5[df_v5['최종발주합계'] > 0].copy()
        if not csv_data.empty:
            st.download_button(
                "📥 발주서 CSV 다운로드", 
                csv_data.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig'), 
                f"발주서_{v5_date.strftime('%m%d')}.csv", 
                use_container_width=True
            )



# ==========================================================
# --- [6단계: 전체 히스토리 관리 (8개 핵심 항목 전용)] ---
# ==========================================================
if st.session_state.get('analyzed'):
    st.divider()
    st.subheader("📜 6단계: 전체 히스토리 관리")

    # 상단 필터 레이아웃
    f1, f2, f3, f4 = st.columns([1, 0.5, 1.2, 1.2])
    with f1:
        today = datetime.now(KST).date()
        d_range = st.date_input("🗓️ 날짜 범위", value=(today, today), key="v6_final_date")
    with f2:
        st.write(""); st.write("")
        search_trigger = st.button("🔍 검색", use_container_width=True, type="primary")
    with f3:
        h_q = st.text_input("🔍 상품명 검색", placeholder="결과 내 검색...", key="v6_final_q")
    with f4:
        # 회차 선택은 검색 후에 데이터가 있으면 업데이트됨
        selected_batch = st.selectbox("📥 저장 회차 선택", ["전체보기"], key="v6_final_batch")

    # 검색 실행
    if search_trigger or h_q:
        try:
            with st.spinner("📡 데이터를 가져오는 중..."):
                sh = get_sheet().worksheet("발주기록")
                vals = sh.get_all_values()
            
            if len(vals) > 1:
                # 8개 규격 컬럼 설정
                target_cols = ["날짜시간", "상품명", "옵션", "공급쳐상품명", "가용재고", "리오더수량", "추가발주수량", "권장 발주수량"]
                df_hist = pd.DataFrame(vals[1:], columns=vals[0] if len(vals[0]) == len(target_cols) else None)
                
                # 만약 컬럼명이 틀어져있어도 강제로 8개 자르기
                if df_hist.shape[1] >= 8:
                    df_hist = df_hist.iloc[:, :8]
                    df_hist.columns = target_cols

                # 날짜 필터링
                df_hist["날짜_만"] = df_hist["날짜시간"].astype(str).str.slice(0, 10)
                if len(d_range) == 2:
                    s_d, e_d = d_range[0].strftime('%Y-%m-%d'), d_range[1].strftime('%Y-%m-%d')
                    df_hist = df_hist[(df_hist["날짜_만"] >= s_d) & (df_hist["날짜_만"] <= e_d)]
                
                # 상품명 검색
                if h_q:
                    df_hist = df_hist[df_hist["상품명"].str.contains(h_q, case=False)]

                if not df_hist.empty:
                    df_hist = df_hist.sort_values(by="날짜시간", ascending=False)
                    
                    st.success(f"✅ {len(df_hist)}건의 기록을 찾았습니다.")
                    
                    # [표 출력]
                    st.dataframe(
                        df_hist.drop(columns=["날짜_만"]), 
                        use_container_width=True, 
                        hide_index=True,
                        column_config={
                            "상품명": st.column_config.TextColumn("상품명", width=300),
                            "날짜시간": st.column_config.TextColumn("날짜시간", width=180),
                        }
                    )
                    
                    csv = df_hist.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
                    st.download_button("📥 내역 다운로드(CSV)", csv, f"발주기록_{today}.csv", use_container_width=True)
                else:
                    st.warning("🧐 해당 조건에 맞는 기록이 없습니다.")
            else:
                st.info("💡 아직 저장된 기록이 없습니다.")
        except Exception as e:
            st.error(f"📡 오류 발생: {e}")



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
