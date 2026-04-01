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
# [4단계: 데이터 편집 및 재고 관리] - 사장님 지정 컬럼 순서 및 로직 통합
# ------------------------------------------------------------------
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📊 4단계: 데이터 편집 및 재고 관리")

    p = st.session_state.p
    sold_out_col, item, option = p['so'], p['it'], p['op']
    vendor, v_item = p['vn'], p['vi']
    stock, avail, t3day, t7day = p['st'], p['av'], p['t3'], p['t7']
    lt, ss = p['lt'], p['ss']

    # --- [1] 상단 UI (필터 / 검색 / 날짜 선택) ---
    f_c1, f_c2, f_c3 = st.columns([1, 2, 1])
    with f_c1: 
        filter_mode = st.selectbox("🚦 상태 필터", ["전체보기", "정상만", "품절만"], index=1, key="v4_main_filter")
    with f_c2: 
        search_query = st.text_input("🔍 상품 검색", placeholder="상품명 또는 옵션 입력...", key="v4_main_search")
    with f_c3: 
        selected_date = st.date_input("🗓️ 입고 조회/기록 날짜", datetime.now(KST).date(), key="v4_main_date")

    # --- [2] 데이터 로드 로직 (실시간 리오더 합계 & 선택 날짜 입고량) ---
    def get_sheet_data_v4(target_date):
        try:
            # A. 리오더 총합 계산 (발주기록 시트: 기존리order + 추가발주)
            ws_v7 = get_sheet().worksheet("발주기록")
            d7 = ws_v7.get_all_values()
            reorder_map = {}
            if len(d7) > 1:
                df7 = pd.DataFrame(d7[1:], columns=[c.strip() for c in d7[0]])
                c_old = '기존리order'; c_add = '추가발주'
                df7['sum'] = pd.to_numeric(df7[c_old], errors='coerce').fillna(0) + pd.to_numeric(df7[c_add], errors='coerce').fillna(0)
                reorder_map = df7.groupby(['상품명', '옵션'])['sum'].sum().to_dict()

            # B. 선택한 날짜의 과거 입고량 (입고기록 시트)
            ws_h = get_sheet().worksheet("입고기록")
            dh = ws_h.get_all_values()
            history_map = {}
            if len(dh) > 1:
                dfh = pd.DataFrame(dh[1:], columns=[c.strip() for c in dh[0]])
                target_str = target_date.strftime('%Y-%m-%d')
                dfh_filtered = dfh[dfh['날짜'].astype(str).str.contains(target_str)]
                if not dfh_filtered.empty:
                    dfh_filtered['수량'] = pd.to_numeric(dfh_filtered['수량'], errors='coerce').fillna(0)
                    history_map = dfh_filtered.groupby(['상품명', '옵션'])['수량'].sum().to_dict()
            
            return reorder_map, history_map
        except: return {}, {}

    reorder_map, history_map = get_sheet_data_v4(selected_date)

    # --- [3] 데이터 가공 및 계산 ---
    df_work = st.session_state.df_raw.copy()
    for col in [stock, avail, t3day, t7day]:
        if col in df_work.columns:
            df_work[col] = pd.to_numeric(df_work[col], errors='coerce').fillna(0).astype(int)

    # 매핑 및 수치 계산
    df_work["리오더 총합"] = df_work.apply(lambda r: int(reorder_map.get((str(r[item]).strip(), str(r[option]).strip()), 0)), axis=1)
    df_work["과거 입고"] = df_work.apply(lambda r: int(history_map.get((str(r[item]).strip(), str(r[option]).strip()), 0)), axis=1)
    df_work["리오더 입고"] = 0 
    
    # 판매량 및 권장수량 계산
    df_work['일판매'] = df_work.apply(lambda r: int(round(r[t7day]/7)) if r[t7day]>0 else (int(round(r[t3day]/3)) if r[t3day]>0 else 0), axis=1)
    df_work['3일판매'] = (df_work['일판매'] * 3).astype(int)
    df_work['권장수량'] = ((df_work['일판매'] * (lt + ss)) - (df_work[avail] + df_work['리오더 총합'])).clip(lower=0).astype(int)

    # 필터링
    if search_query:
        df_f = df_work[df_work[item].astype(str).str.contains(search_query, case=False) | df_work[option].astype(str).str.contains(search_query, case=False)]
    else:
        is_so = df_work[sold_out_col].astype(str).str.contains('품절', na=False)
        df_f = df_work[~is_so] if filter_mode == "정상만" else (df_work[is_so] if filter_mode == "품절만" else df_work)

    # --- [4] 컬럼 순서 재배치 (사장님 요청 순서) ---
    # 상태, 공급쳐, 상품명, 옵션, 공급처상품명 이후 순서 고정
    df_disp = df_f.rename(columns={sold_out_col: "상태", vendor: "공급쳐", v_item: "공급처상품명", item: "상품명", option: "옵션", stock: "정상", avail: "가용"})
    
    # 사장님이 말씀하신 순서: 정상 -> 가용 -> 리오더총합 -> 리오더입고 -> 과거 입고 -> 3일판매 -> 일판매 -> 권장수량
    final_cols = ["상태", "공급쳐", "상품명", "옵션", "공급처상품명", "정상", "가용", "리오더 총합", "리오더 입고", "과거 입고", "3일판매", "일판매", "권장수량"]

    with st.form("v4_ordered_form"):
        edited_df = st.data_editor(
            df_disp[final_cols], 
            use_container_width=True, 
            hide_index=True, 
            key="v4_editor_final",
            column_config={
                "상품명": st.column_config.TextColumn(width=250),
                "리오더 총합": st.column_config.NumberColumn("리오더 총합", disabled=True, help="시트 합계 (수정불가)"),
                "과거 입고": st.column_config.NumberColumn(f"{selected_date.strftime('%m/%d')} 입고", disabled=True),
                "리오더 입고": st.column_config.NumberColumn("리오더입고", min_value=0),
                "3일판매": st.column_config.NumberColumn("3일판매", disabled=True),
                "일판매": st.column_config.NumberColumn("일판매", disabled=True),
                "권장수량": st.column_config.NumberColumn("권장수량", disabled=True)
            }
        )
        
        if st.form_submit_button("💾 입고 정보 저장 및 리오더 차감", use_container_width=True, type="primary"):
            edits = st.session_state["v4_editor_final"].get("edited_rows", {})
            if edits:
                try:
                    v7_sh, h_sh = get_sheet().worksheet("발주기록"), get_sheet().worksheet("입고기록")
                    for r_idx, changes in edits.items():
                        actual_idx = df_disp.index[int(r_idx)]
                        if "리오더 입고" in changes:
                            qty = int(changes["리오더 입고"])
                            if qty > 0:
                                # 1. 발주기록 시트 '추가발주' 열에 마이너스(-) 한 줄 추가
                                v7_sh.append_row([datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S'), str(df_disp.at[actual_idx, "상품명"]), str(df_disp.at[actual_idx, "옵션"]), str(df_disp.at[actual_idx, "공급처상품명"]), "0", "0", -qty, "0", "4단계 입고반영", "시스템"])
                                # 2. 입고기록 시트 저장 (선택한 날짜 기준)
                                h_sh.append_row([selected_date.strftime('%Y-%m-%d'), str(df_disp.at[actual_idx, "상품명"]), str(df_disp.at[actual_idx, "옵션"]), qty])
                    st.success("✅ 입고 정보가 성공적으로 반영되었습니다!"); time.sleep(0.5); st.rerun()
                except Exception as e: st.error(f"❌ 저장 중 오류 발생: {e}")



# ------------------------------------------------------------------
# [5단계: 최종 발주 요약 및 구글 시트 저장] - 사장님 맞춤 순서 및 에러 수정판
# ------------------------------------------------------------------
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📋 5단계: 최종 발주 리스트 요약")

    # 1. 4단계 가공 데이터(df_work) 활용 (품절 제외)
    # df_work는 4단계에서 '리오더 총합', '권장수량' 등이 계산된 상태여야 합니다.
    df_v5 = df_work[~df_work[sold_out_col].astype(str).str.contains('품절', na=False)].copy()
    
    # 상단 필터/검색 UI
    f1, f2, f3 = st.columns([1.5, 2, 1])
    with f1:
        m5_f = st.selectbox("🚦 상태 필터", ["🚨 고위험/주의", "✅ 전체정상"], key="v5_main_filter")
    with f2:
        s5_q = st.text_input("🔍 상품명/옵션 검색", key="v5_main_search", placeholder="검색어를 입력하세요...")
    with f3:
        d5_d = st.date_input("🗓️ 발주 기록 날짜", datetime.now(KST).date(), key="v5_main_date")

    # 추가발주 임시 저장소 확인
    if 'add_order_dict' not in st.session_state: 
        st.session_state.add_order_dict = {}
    
    # 2. 데이터 전처리 (에디터 표시용)
    df_v5['추가발주입력'] = df_v5.index.map(st.session_state.add_order_dict).fillna(0).astype(int)
    df_v5['상태표시'] = df_v5['권장수량'].apply(lambda x: "🚨 긴급" if x > 0 else "✅ 정상")

    # 3. 검색 및 필터 적용
    if s5_q:
        df_v5_view = df_v5[df_v5[item].astype(str).str.contains(s5_q, case=False) | 
                           df_v5[option].astype(str).str.contains(s5_q, case=False)].copy()
    else:
        # 고위험군은 권장수량이 0보다 큰 상품 기준
        danger_names = df_v5[df_v5['권장수량'] > 0][item].unique() if '권장수량' in df_v5.columns else []
        if m5_f == "🚨 고위험/주의":
            df_v5_view = df_v5[df_v5[item].isin(danger_names)].copy()
        else:
            df_v5_view = df_v5[~df_v5[item].isin(danger_names)].copy()
    
    df_v5_view = df_v5_view.sort_values(by=[item, option])

    # 4. [중요] 사장님 요청 셀 순서 매핑
    # 원본컬럼명 : 화면표시컬럼명
    view_map = {
        "상태표시": "상태", 
        item: "상품명", 
        option: "옵션", 
        v_item: "공급처상품명", 
        "리오더 총합": "리오더총합", 
        "추가발주입력": "추가발주", 
        "권장수량": "권장수량",
        "비고": "메모"
    }
    
    # KeyError 방지를 위해 실제 존재하는 컬럼만 필터링
    valid_keys = [k for k in view_map.keys() if k in df_v5_view.columns]
    
    # 5. 에디터 폼 출력
    with st.form("v5_master_form"):
        # 요청하신 순서대로 컬럼 재배치 및 이름 변경
        df_ed_v5 = df_v5_view[valid_keys].rename(columns=view_map)
        
        st.data_editor(
            df_ed_v5, 
            use_container_width=True, 
            hide_index=True, 
            key="v5_editor",
            column_config={
                "상태": st.column_config.TextColumn(width=70),
                "상품명": st.column_config.TextColumn(width=200),
                "옵션": st.column_config.TextColumn(width=150),
                "공급처상품명": st.column_config.TextColumn(width=200),
                "리오더총합": st.column_config.NumberColumn(disabled=True, width=90),
                "추가발주": st.column_config.NumberColumn("추가발주", min_value=0, width=90), # 입력가능
                "권장수량": st.column_config.NumberColumn(disabled=True, width=90),
                "메모": st.column_config.TextColumn("메모(비고)", width=300) # 입력가능
            },
            # 추가발주와 메모만 수정할 수 있게 설정
            disabled=[col for col in df_ed_v5.columns if col not in ["추가발주", "메모"]]
        )
        
        if st.form_submit_button("✅ 1. 추가발주 및 메모 확정 (기록 준비)", use_container_width=True, type="primary"):
            edits_v5 = st.session_state["v5_editor"].get("edited_rows", {})
            if edits_v5:
                for r_idx, val in edits_v5.items():
                    # 에디터 행 번호를 데이터프레임 인덱스로 매칭
                    idx = df_v5_view.index[int(r_idx)]
                    if "추가발주" in val:
                        st.session_state.add_order_dict[idx] = int(val["추가발주"])
                    if "메모" in val:
                        # 원본 데이터(df_raw)에 메모(비고) 내용 저장
                        st.session_state.df_raw.at[idx, "비고"] = str(val["메모"])
                st.success("✅ 변경 내용이 확정되었습니다! 아래 저장 버튼을 눌러주세요."); time.sleep(0.5); st.rerun()
            else:
                st.info("💡 수정된 내용이 없습니다.")

    st.divider()
    
    # 6. 저장 및 다운로드 버튼 섹션
    c_save, c_down = st.columns(2)
    
    with c_save:
        if st.button("💾 2. 구글 시트에 발주 기록 최종 저장", use_container_width=True, type="primary"):
            # 추가 수량이 0보다 큰 항목들만 추출
            valid_ids = [k for k, v in st.session_state.add_order_dict.items() if v > 0]
            if not valid_ids:
                st.warning("⚠️ 저장할 추가 수량이 없습니다. [추가발주]를 입력하고 확정 버튼을 먼저 눌러주세요.")
            else:
                try:
                    with st.spinner("🚀 발주기록 시트로 전송 중..."):
                        ws_log = get_sheet().worksheet("발주기록")
                        now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
                        log_rows = []
                        for idx in valid_ids:
                            row = st.session_state.df_raw.loc[idx]
                            add_qty = st.session_state.add_order_dict[idx]
                            
                            # 시트 컬럼 순서: 발주시간, 상품명, 옵션, 공급처상품명, 기존리order(총합), 추가발주, 메모, 업체명
                            log_rows.append([
                                now_s, 
                                str(row.get(item, "")), 
                                str(row.get(option, "")), 
                                str(row.get(v_item, "")), 
                                int(df_v5.at[idx, '리오더 총합']), 
                                int(add_qty), 
                                str(row.get('비고', "")), 
                                str(row.get(vendor, ""))
                            ])
                        
                        if log_rows:
                            ws_log.append_rows(log_rows)
                            st.session_state.add_order_dict = {} # 저장 후 딕셔너리 초기화
                            st.success("✅ 구글 시트 저장 성공!"); time.sleep(1); st.rerun()
                except Exception as e: 
                    st.error(f"❌ 저장 중 오류 발생: {e}")

    with c_down:
        # CSV 다운로드용 (추가발주 수량이 있는 것만)
        csv_target = df_v5[df_v5.index.isin(st.session_state.add_order_dict.keys())].copy()
        csv_target['최종발주'] = csv_target.index.map(st.session_state.add_order_dict)
        csv_target = csv_target[csv_target['최종발주'] > 0]
        
        if not csv_target.empty:
            csv_res = csv_target[[vendor, item, option, v_item, '최종발주']].rename(columns={
                vendor: "공급처", item: "상품명", option: "옵션", v_item: "공급처상품명", "최종발주": "발주수량"
            })
            csv_data = csv_res.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button("📥 최종 발주서(CSV) 다운로드", data=csv_data, file_name=f"발주서_{d5_d.strftime('%m%d')}.csv", mime="text/csv", use_container_width=True)
        else:
            st.button("📥 다운로드 (데이터 없음)", disabled=True, use_container_width=True)





# ------------------------------------------------------------------
# [6단계: 전체 히스토리 관리] - 4/5단계 변경 시트 구조 반영 버전
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.subheader("📜 6단계: 전체 히스토리 관리")

    f1, f2, f3, f4 = st.columns([1.2, 0.8, 1.5, 1.5])
    
    with f1:
        today = datetime.now(KST).date()
        # 조회 범위 선택 (기본값: 오늘)
        d_range = st.date_input("🗓️ 1. 조회 범위", value=(today, today), key="v6_date_range")
    
    with f2:
        st.write(""); st.write("") 
        search_trigger = st.button("🔍 2. 내역 조회", use_container_width=True, type="primary")

    # 세션 상태 초기화
    if 'v6_data' not in st.session_state: st.session_state.v6_data = None
    if 'v6_sessions' not in st.session_state: st.session_state.v6_sessions = []
    if 'v6_display_text' not in st.session_state: st.session_state.v6_display_text = ""

    # [내역 조회 로직]
    if search_trigger:
        try:
            with st.spinner("📡 발주 내역을 불러오는 중..."):
                worksheet = get_sheet().worksheet("발주기록")
                all_h = worksheet.get_all_values()
                if len(all_h) > 1:
                    df_all = pd.DataFrame(all_h[1:])
                    # 🚨 사장님 시트 컬럼 순서에 맞춤 (8개 핵심 컬럼)
                    df_all.columns = ["발주시간", "상품명", "옵션", "공급처상품명", "기존리order", "추가발주", "메모", "업체명"]
                    
                    # 날짜 필터링을 위한 임시 컬럼
                    df_all["날짜_만"] = df_all["발주시간"].astype(str).str.slice(0, 10)
                    
                    # 날짜 범위 처리
                    curr_str = datetime.now(KST).strftime('%Y-%m-%d')
                    if isinstance(d_range, (list, tuple)) and len(d_range) == 2:
                        s_d, e_d = d_range[0].strftime('%Y-%m-%d'), d_range[1].strftime('%Y-%m-%d')
                        st.session_state.v6_display_text = f"🗓️ {s_d} ~ {e_d}"
                    elif isinstance(d_range, (list, tuple)) and len(d_range) == 1:
                        s_d, e_d = d_range[0].strftime('%Y-%m-%d'), curr_str
                        st.session_state.v6_display_text = f"🗓️ {s_d} ~ {e_d} (오늘까지)"
                    else:
                        s_d, e_d = "0000-00-00", "9999-99-99"
                        st.session_state.v6_display_text = "🗓️ 전체 내역"
                    
                    # 필터링 및 세션(회차) 저장
                    df_filtered = df_all[(df_all["날짜_만"] >= s_d) & (df_all["날짜_만"] <= e_d)].copy()
                    st.session_state.v6_data = df_filtered
                    # 최근 발주가 위로 오도록 정렬하여 세션 리스트 생성
                    st.session_state.v6_sessions = sorted(df_filtered["발주시간"].unique(), reverse=True)
                else:
                    st.session_state.v6_data = None
                    st.info("💡 저장된 내역이 없습니다.")
        except Exception as e:
            st.error(f"📡 데이터를 불러오지 못했습니다: {e}")

    # [필터 및 회차 선택 UI]
    with f3: h_q = st.text_input("🔍 3. 상품명/옵션 검색", key="v6_search_q")
    with f4:
        if st.session_state.v6_sessions:
            session_options = ["📊 선택 범위 전체 합산"] + [f"{len(st.session_state.v6_sessions)-i}회차 ({t[5:16]})" for i, t in enumerate(st.session_state.v6_sessions)]
            sel_session_label = st.selectbox("📦 4. 회차 선택", session_options, key="v6_session_select")
        else:
            st.selectbox("📦 4. 회차 선택", ["조회 결과 없음"], disabled=True)
            sel_session_label = None

    # [데이터 출력창]
    if st.session_state.v6_data is not None and sel_session_label:
        df_display = st.session_state.v6_data.copy()
        
        # 숫자 변환 (합산을 위해 필요)
        df_display["추가발주"] = pd.to_numeric(df_display["추가발주"], errors='coerce').fillna(0)
        df_display["기존리order"] = pd.to_numeric(df_display["기존리order"], errors='coerce').fillna(0)

        # 회차별 필터링
        if sel_session_label == "📊 선택 범위 전체 합산":
            display_title = st.session_state.v6_display_text + " 발주 합계"
            # 전체 합산 로직 (가용/권장 제외)
            df_display = df_display.groupby(["업체명", "상품명", "옵션", "공급처상품명"], as_index=False).agg({
                "발주시간": "max", 
                "기존리order": "last", # 마지막 기준 총합
                "추가발주": "sum",     # 선택 기간 내 추가발주 총합
                "메모": lambda x: " / ".join(set(filter(None, x.astype(str))))
            })
        else:
            target_time = st.session_state.v6_sessions[session_options.index(sel_session_label)-1]
            df_display = df_display[df_display["발주시간"] == target_time].copy()
            display_title = f"✅ {sel_session_label} 상세 내역"

        # 상품명/옵션 검색 적용
        if h_q:
            df_display = df_display[
                df_display["상품명"].astype(str).str.contains(h_q, case=False) | 
                df_display["옵션"].astype(str).str.contains(h_q, case=False)
            ]

        # 최종 화면 출력
        if not df_display.empty:
            st.write(f"#### {display_title}")
            # 표시할 컬럼 순서 (사장님 시트 구조와 일치)
            view_order = ["발주시간", "업체명", "상품명", "옵션", "공급처상품명", "기존리order", "추가발주", "메모"]
            st.dataframe(df_display[view_order], use_container_width=True, hide_index=True)
            
            # CSV 다운로드
            csv_data = df_display[view_order].to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button(
                label=f"📥 {display_title} CSV 다운로드", 
                data=csv_data, 
                file_name=f"발주히스토리_{datetime.now().strftime('%m%d')}.csv", 
                mime="text/csv",
                use_container_width=True
            )



# ------------------------------------------------------------------
# [7단계: 실시간 리오더 최종 잔량 상황판] - 기간 제한 없이 잔량 있는 모든 업체 노출
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.subheader("🚀 7단계: 실시간 리오더 최종 잔량 상황판")

    try:
        ws_v7 = get_sheet().worksheet("발주기록")
        all_v7 = ws_v7.get_all_values()
        
        if len(all_v7) > 1:
            header_v7 = [c.strip() for c in all_v7[0]]
            df_raw_v7 = pd.DataFrame(all_v7[1:], columns=header_v7)
            
            # 1. 데이터 전처리 (공백 제거 및 숫자 변환)
            if "업체명" in df_raw_v7.columns:
                df_raw_v7["업체명"] = df_raw_v7["업체명"].astype(str).str.strip()
            
            qty_col = '추가발주' if '추가발주' in df_raw_v7.columns else ('추가' if '추가' in df_raw_v7.columns else df_raw_v7.columns[5])
            df_raw_v7[qty_col] = pd.to_numeric(df_raw_v7[qty_col], errors='coerce').fillna(0).astype(int)
            
            # 2. 상단 필터 영역 (순서 유지)
            f1, f2, f3, f4 = st.columns([1.2, 0.8, 1.5, 1.5])
            with f1:
                # 💡 시작 날짜를 아예 '전체'로 인식하게끔 아주 과거로 기본값 설정
                default_start = datetime(2025, 1, 1).date() 
                d_range_v7 = st.date_input("🗓️ 1. 조회 시작일 (과거 발주분 포함)", value=(default_start, datetime.now(KST).date()), key="v7_range")
            with f2:
                st.write(""); st.write("")
                if st.button("📈 2. 데이터 동기화", use_container_width=True, type="primary"):
                    st.rerun()
            with f3:
                q_v7 = st.text_input("🔍 3. 상품/옵션 검색", key="v7_search_q")
            with f4:
                v_list = sorted(df_raw_v7["업체명"].unique().tolist())
                v_choice = st.selectbox("🏭 4. 업체 선택", ["전체 업체"] + v_list, key="v7_vendor_sel")

            # --- [필터링 로직 수정: 날짜는 참고용, 잔량 위주로] ---
            if isinstance(d_range_v7, (list, tuple)) and len(d_range_v7) == 2:
                s_d, e_d = d_range_v7[0].strftime('%Y-%m-%d'), d_range_v7[1].strftime('%Y-%m-%d')
            else:
                s_d = d_range_v7[0].strftime('%Y-%m-%d') if isinstance(d_range_v7, (list, tuple)) else d_range_v7.strftime('%Y-%m-%d')
                e_d = datetime.now(KST).strftime('%Y-%m-%d')

            # 일단 전체 데이터를 가져온 뒤 그룹화 (날짜 필터링을 그룹화 '후'에 하거나 범위를 넓게 잡음)
            df_all = df_raw_v7.copy()
            
            # 업체/상품/옵션별 최종 잔량 계산 (날짜 상관없이 전체 합산)
            group_cols = ["업체명", "상품명", "옵션", "공급처상품명"]
            df_total_balance = df_all.groupby(group_cols, as_index=False).agg({
                qty_col: "sum",
                "메모": lambda x: " / ".join(set(filter(None, x.astype(str))))
            }).rename(columns={qty_col: "미입고 잔량"})

            # 🚨 잔량이 0보다 큰 것만 남김 (러블리마켓이 1개라도 남았으면 여기서 걸러짐)
            df_display = df_total_balance[df_total_balance["미입고 잔량"] > 0].copy()

            # 선택한 필터 적용 (업체 선택/검색어)
            if v_choice != "전체 업체":
                df_display = df_display[df_display["업체명"] == v_choice]
            if q_v7:
                df_display = df_display[df_display["상품명"].str.contains(q_v7, case=False) | df_display["옵션"].str.contains(q_v7, case=False)]

            # ---------------------------------------------------------
            # [3. 출력 영역]
            # ---------------------------------------------------------
            if not df_display.empty:
                # 상단 업체별 요약
                df_vendor_sum = df_display.groupby("업체명")["미입고 잔량"].sum().reset_index().sort_values(by="미입고 잔량", ascending=False)
                st.write(f"#### 🏭 업체별 현재 잔량 요약 (총 {df_vendor_sum['미입고 잔량'].sum():,}개)")
                v_cols = st.columns(4)
                for i, r in df_vendor_sum.reset_index(drop=True).iterrows():
                    with v_cols[i % 4]:
                        st.metric(label=r["업체명"], value=f"{r['미입고 잔량']:,}개")
                st.divider()

                # 하단 상세 리스트 (날짜 열은 가장 최근 기록 날짜를 보여주도록 추가 처리 가능하지만, 우선 순서대로 출력)
                st.write(f"#### 📋 상세 미입고 리스트 (잔량 있는 품목 전체)")
                # 화면 표시 순서: 업체명 -> 상품명 -> 옵션 -> 공급처상품명 -> 미입고 잔량 -> 메모
                display_cols = ["업체명", "상품명", "옵션", "공급처상품명", "미입고 잔량", "메모"]
                st.dataframe(df_display[display_cols], use_container_width=True, hide_index=True)
            else:
                st.info("🔎 현재 잔량이 남은 데이터가 없습니다. 모든 리오더가 입고 완료되었거나 검색 조건에 맞는 데이터가 없습니다.")

        else:
            st.info("💡 발주 데이터가 없습니다.")
            
    except Exception as e:
        st.error(f"📡 상황판 로딩 오류: {e}")
