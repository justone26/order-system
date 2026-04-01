import streamlit as st
import pandas as pd
import numpy as np
import time
import io
import streamlit.components.v1 as components
from datetime import datetime, timedelta, timezone

# 🚨 한국 시간 설정
KST = timezone(timedelta(hours=9)) 
current_today = datetime.now(KST).date()

# --- [세션 상태 초기화] ---
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'analyzed' not in st.session_state: st.session_state.analyzed = False
if 'p' not in st.session_state: st.session_state.p = {}
if 'add_order_dict' not in st.session_state: st.session_state.add_order_dict = {}
if 'upload_key' not in st.session_state: st.session_state.upload_key = 0

# --- [새로고침 방지] ---
components.html("<script>window.onbeforeunload = function() { return '변경사항이 저장되지 않을 수 있습니다.'; };</script>", height=0)

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

# --- [핵심: 리오더 동기화 및 진단] ---
def sync_reorder_from_sheet(df_uploaded):
    try:
        sh = get_sheet()
        if not sh:
            st.error("❌ 구글 시트 연결 실패 (Secrets 설정 확인 필요)")
            return df_uploaded
        
        ws = sh.worksheet("발주기록")
        all_data = ws.get_all_values()
        
        if len(all_data) <= 1:
            st.warning("⚠️ '발주기록' 시트에 데이터가 없습니다.")
            return df_uploaded
            
        reorder_map = {}
        # 시트 데이터 루프 (B:상품명, C:옵션, F:기존, G:추가)
        for row in all_data[1:]:
            try:
                # 1. 시트 데이터 표준화 (공백 제거, 대문자)
                s_name = str(row[1]).strip().replace(" ", "").upper()
                s_opt = str(row[2]).strip().replace(" ", "").upper()
                
                if not s_name: continue
                
                # 2. 숫자 변환 (쉼표 제거 및 에러 방지)
                def clean_num(v):
                    try:
                        return int(float(str(v).replace(",", "").strip()))
                    except:
                        return 0
                
                # F열(index 5) + G열(index 6) 합산
                qty = clean_num(row[5]) + clean_num(row[6])
                
                if qty > 0:
                    key = s_name + "_" + s_opt
                    reorder_map[key] = reorder_map.get(key, 0) + qty
            except:
                continue

        # 3. 진단 메시지 출력
        if not reorder_map:
            st.info("ℹ️ 시트에서 수량이 입력된 상품을 찾지 못했습니다. (F, G열 확인)")
        else:
            st.success("✅ 시트에서 총 " + str(len(reorder_map)) + "건의 리오더 품목을 로드했습니다.")

        # 4. 업로드 파일 매칭
        if "리오더 수량" in df_uploaded.columns:
            df_uploaded = df_uploaded.drop(columns=["리오더 수량"])

        def get_final_qty(r):
            # 업로드 파일의 상품명/옵션도 공백 제거 후 비교
            u_key = (str(r['상품명']).strip().replace(" ", "") + "_" + 
                     str(r['옵션']).strip().replace(" ", "")).upper()
            return reorder_map.get(u_key, 0)

        df_uploaded['리오더 수량'] = df_uploaded.apply(get_final_qty, axis=1)
        
        return df_uploaded

    except Exception as e:
        # 🚨 f-string 오류 방지를 위해 일반 결합 사용
        st.error("시스템 오류 발생: " + str(e))
        return df_uploaded

# --- [보조 함수] ---
def get_incoming_history():
    try:
        sh = get_sheet() 
        ws = sh.worksheet("입고기록")
        data = ws.get_all_records()
        if not data: return pd.DataFrame(columns=['상품명', '옵션', '과거리오더 입고'])
        df_h = pd.DataFrame(data)
        df_h.columns = [c.strip() for c in df_h.columns]
        df_h['상품명'] = df_h['상품명'].astype(str).str.strip()
        df_h['옵션'] = df_h['옵션'].astype(str).str.strip()
        summary = df_h.groupby(['상품명', '옵션'])['수량'].sum().reset_index()
        summary.rename(columns={'수량': '과거리오더 입고'}, inplace=True)
        return summary
    except: 
        return pd.DataFrame(columns=['상품명', '옵션', '과거리오더 입고'])
        

# [0] 공통 함수 (4단계와 5단계가 같이 쓰도록 밖으로 뺐습니다)
def get_realtime_data_v4(target_date):
    try:
        import unicodedata
        import re
        def super_clean(t):
            t = unicodedata.normalize('NFC', str(t))
            return re.sub(r'[^a-zA-Z0-9가-힣]', '', t).upper()

        ws_v7 = get_sheet().worksheet("발주기록")
        d7 = ws_v7.get_all_values()
        r_map = {}
        if len(d7) > 1:
            for row in d7[1:]:
                try:
                    s_name = super_clean(row[1])
                    s_opt = super_clean(row[2])
                    if not s_name: continue
                    def to_i(v):
                        try: return int(float(str(v).replace(",", "").strip()))
                        except: return 0
                    # A~J열 중 리오더 수치 합산
                    total_qty = to_i(row[5]) + to_i(row[6])
                    if total_qty != 0:
                        key = s_name + "_" + s_opt
                        r_map[key] = r_map.get(key, 0) + total_qty
                except: continue

        ws_h = get_sheet().worksheet("입고기록")
        dh = ws_h.get_all_values()
        h_map = {}
        if len(dh) > 1:
            t_str = target_date.strftime('%Y-%m-%d')
            for row_h in dh[1:]:
                try:
                    if t_str in str(row_h[0]):
                        h_key = super_clean(row_h[1]) + "_" + super_clean(row_h[2])
                        h_map[h_key] = h_map.get(h_key, 0) + int(float(str(row_h[3]).replace(",", "")))
                except: continue
        return r_map, h_map
    except Exception as e:
        return {}, {}



def sync_reorder_from_sheet(df_uploaded):
    try:
        sh = get_sheet()
        if not sh: return df_uploaded
        ws = sh.worksheet("발주기록")
        all_data = ws.get_all_values()
        if len(all_data) <= 1: return df_uploaded
            
        # 1. 헤더 위치 자동 찾기 (날짜 칸이 어디 있든 상관없음)
        header = [str(h).strip().replace(" ", "") for h in all_data[0]]
        idx_name = next((i for i, h in enumerate(header) if "상품명" in h), 1)
        idx_opt = next((i for i, h in enumerate(header) if "옵션" in h), 2)
        idx_f = next((i for i, h in enumerate(header) if "기존" in h), 5)
        idx_g = next((i for i, h in enumerate(header) if "추가" in h), 6)

        reorder_map = {}
        import unicodedata
        import re

        # 2. 시트 데이터 정리 (특수문자 싹 제거하고 알맹이 글자만 추출)
        for row in all_data[1:]:
            try:
                def super_clean(t):
                    t = unicodedata.normalize('NFC', str(t))
                    return re.sub(r'[^a-zA-Z0-9가-힣]', '', t).upper()

                s_name = super_clean(row[idx_name])
                s_opt = super_clean(row[idx_opt])
                if not s_name: continue
                
                def to_i(v):
                    try: return int(float(str(v).replace(",", "").strip()))
                    except: return 0
                
                qty = to_i(row[idx_f]) + to_i(row[idx_g])
                if qty > 0:
                    key = s_name + "_" + s_opt
                    reorder_map[key] = reorder_map.get(key, 0) + qty
            except:
                continue

        # 3. 업로드된 엑셀과 매칭 (글자가 포함만 되어도 매칭)
        if "리오더 수량" in df_uploaded.columns:
            df_uploaded = df_uploaded.drop(columns=["리오더 수량"])

        def final_match(r):
            u_name = re.sub(r'[^a-zA-Z0-9가-힣]', '', str(r['상품명'])).upper()
            u_opt = re.sub(r'[^a-zA-Z0-9가-힣]', '', str(r['옵션'])).upper()
            u_key = u_name + "_" + u_opt
            
            # 1순위: 완벽 일치 / 2순위: 포함 관계 확인
            if u_key in reorder_map: return reorder_map[u_key]
            for k, v in reorder_map.items():
                if u_name in k and u_opt in k: return v
            return 0

        df_uploaded['리오더 수량'] = df_uploaded.apply(final_match, axis=1)
        
        # 결과 리포트
        matched_cnt = (df_uploaded['리오더 수량'] > 0).sum()
        if matched_cnt > 0:
            st.success(f"✅ 날짜 칸 무시 성공! {matched_cnt}건의 수량을 매칭했습니다.")
        else:
            st.warning("⚠️ 시트에서 101개를 읽었으나 엑셀과 이름이 달라 매칭 실패했습니다.")
            
        return df_uploaded
    except Exception as e:
        st.error(f"매칭 오류: {str(e)}")
        return df_uploaded

        # 3. 매칭 데이터 생성
        reorder_map = {}
        for i, row in enumerate(all_data[1:]):
            # 상품명(B), 옵션(C), 기존리오더(F), 추가발주(G)
            name = str(row[1]).strip().replace(" ", "").upper()
            opt = str(row[2]).strip().replace(" ", "").upper()
            
            def to_int(v):
                try: return int(float(str(v).replace(",", "")))
                except: return 0
            
            qty = to_int(row[5]) + to_int(row[6])
            if qty > 0:
                key = f"{name}_{opt}"
                reorder_map[key] = reorder_map.get(key, 0) + qty

        # 4. 결과 보고
        if not reorder_map:
            st.info("ℹ️ [진단] 시트에서 수량이 0보다 큰 상품을 하나도 찾지 못했습니다. F, G열을 확인하세요.")
        else:
            st.success(f"✅ [진단] 시트에서 총 {len(reorder_map)}개의 리오더 품목을 읽어왔습니다.")

        # 5. 업로드 파일과 매칭
        if "리오더 수량" in df_uploaded.columns:
            df_uploaded = df_uploaded.drop(columns=["리오더 수량"])

        def match_val(r):
            u_key = (str(r['상품명']).strip().replace(" ", "") + "_" + 
                     str(r['옵션']).strip().replace(" ", "")).upper()
            return reorder_map.get(u_key, 0)

        df_uploaded['리오더 수량'] = df_uploaded.apply(match_val, axis=1)
        return df_uploaded

    except Exception as e:
        st.error(f"🔥 [치명적 오류] 시스템 에러 발생: {e}")
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
    
    # 파일 업로드 위젯
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

    # --- 데이터 로드 및 시트 동기화 로직 ---
    if up_file:
        # 파일이 처음 올라왔거나, 아직 데이터프레임이 생성되지 않았을 때 실행
        if st.session_state.get('df_raw') is None:
            try:
                # 1. 파일 읽기
                df = pd.read_csv(up_file) if up_file.name.endswith('.csv') else pd.read_excel(up_file)
                df.columns = [str(c).strip() for c in df.columns] # 컬럼명 공백 제거
                
                # 2. ⭐ [핵심] 구글 시트에서 리오더 수량 가져오기 (날짜 칸 무시 로직 포함된 함수)
                with st.spinner("🔄 구글 시트(발주기록)에서 리오더 수량을 실시간 매칭 중..."):
                    # 여기서 우리가 고친 sync_reorder_from_sheet 함수를 실행합니다.
                    df = sync_reorder_from_sheet(df)
                
                # 3. 만약 리오더 수량 컬럼이 끝까지 안 생겼다면 0으로 채워줌
                if "리오더 수량" not in df.columns: 
                    df["리오더 수량"] = 0
                
                df = df.fillna("") 
                st.session_state.df_raw = df
                # 성공 메시지는 sync_reorder_from_sheet 내부에서 띄워줌
                
            except Exception as e:
                st.error(f"파일 로드 오류: {e}")

    # --- 2~3단계: 매핑 및 분석 설정 (데이터가 로드된 경우에만 표시) ---
    if st.session_state.get('df_raw') is not None:
        st.divider()
        
        # --- 2단계: 매핑 항목 ---
        st.subheader("📋 2단계: 매핑 항목")
        st.info("💡 '리오더 수량'은 시트에서 자동으로 가져왔습니다. 나머지 항목을 확인해주세요.")
        
        cols = st.session_state.df_raw.columns.tolist()
        
        def auto_idx(keys, exclude_keys=None):
            for i, c in enumerate(cols):
                column_name = str(c)
                if exclude_keys and any(ek in column_name for ek in exclude_keys): continue
                if any(k in column_name for k in keys): return i
            return 0

        # 5개씩 2열로 배치
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
            
            t3_target = "3일 발주합계"
            t3_idx = cols.index(t3_target) if t3_target in cols else auto_idx(['3일'], exclude_keys=['1주', '7일', '품절'])
            t3 = st.selectbox("🔥 3일 판매", cols, index=t3_idx, key="sel_t3")
            
            t7_target = "1주발주합계"
            t7_idx = cols.index(t7_target) if t7_target in cols else auto_idx(['7일', '1주'], exclude_keys=['3일', '품절'])
            t7 = st.selectbox("📅 7일 판매", cols, index=t7_idx, key="sel_t7")
            
            reg = st.selectbox("📆 상품 등록일", cols, index=auto_idx(['등록일', '등록일자', '최초등록']), key="sel_reg")

      # --- 3단계: 데이터 분석 설정 ---
    st.subheader("🚀 3단계: 데이터 분석 설정")
    s1, s2 = st.columns(2)
    with s1:
        lt_val = st.number_input("⏳ 리드타임 (일)", value=7, key="inp_lt")
    with s2:
        ss_val = st.number_input("🛡️ 안전재고 (일)", value=3, key="inp_ss")

    if st.button("📊 데이터 분석 시작", use_container_width=True, type="primary"):
        st.session_state.p = {
            'so': so, 'vn': vn, 'vi': vi, 'it': it, 'op': op, 
            'st': stk, 'av': av, 't3': t3, 't7': t7, 'reg': reg,
            'lt': lt_val, 'ss': ss_val
        }
        
        # [수정 포인트] 시트에서 가져온 '리오더 수량'이 포함된 원본 데이터를 복사합니다.
        df_final = st.session_state.df_raw.copy()
        
        # 등록일 날짜 형식 변환
        if reg in df_final.columns:
            df_final[reg] = pd.to_datetime(df_final[reg], errors='coerce')
        
        # ⭐ 중요: 리오더 수량이 숫자인지 확인하고 빈칸은 0으로 채웁니다.
        if "리오더 수량" in df_final.columns:
            df_final["리오더 수량"] = pd.to_numeric(df_final["리오더 수량"], errors='coerce').fillna(0)
        else:
            df_final["리오더 수량"] = 0

        # 수정된 데이터를 세션에 다시 저장하고 화면을 새로고침합니다.
        st.session_state.df_raw = df_final 
        st.session_state.analyzed = True   
        st.rerun()



# ------------------------------------------------------------------
# [4단계: 데이터 편집 및 재고 관리] - 시트 연결(NoneType) 오류 방지 로직 추가
# ------------------------------------------------------------------
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📊 4단계: 데이터 편집 및 재고 관리")

    # ⭐ [핵심 추가] 시트 연결이 끊겼는지 확인하고 다시 불러옵니다.
    try:
        sh = get_sheet()
        if sh is None:
            st.error("❗ 구글 시트 연결에 실패했습니다. 상단의 인증 버튼을 다시 눌러주시거나 새로고침 해주세요.")
            st.stop()
    except NameError:
        st.error("❗ get_sheet() 함수를 찾을 수 없습니다. 코드 상단에 인증 로직이 있는지 확인해주세요.")
        st.stop()

    # [1] 설정값 로드
    p = st.session_state.p
    s_out, item, opt = p['so'], p['it'], p['op']
    vnd, v_it = p['vn'], p['vi']
    stk, avl, t3, t7 = p['st'], p['av'], p['t3'], p['t7']
    lt, ss = p['lt'], p['ss']

    # [2] UI (필터/검색)
    c1, c2, c3 = st.columns([1, 2, 1])
    with c1: f_mode = st.selectbox("🚦 상태 필터", ["전체보기", "정상만", "품절만"], index=1, key="v4_f")
    with c2: s_query = st.text_input("🔍 상품 검색", key="v4_s")
    with c3: s_date = st.date_input("🗓️ 입고 조회 날짜", datetime.now(KST).date(), key="v4_d")

    # [3] 데이터 계산 (get_realtime_data_v4 호출)
    reorder_map, history_map = get_realtime_data_v4(s_date)
    df_work = st.session_state.df_raw.copy()

    # 숫자 형식 변환
    for col in [stk, avl, t3, t7]:
        df_work[col] = pd.to_numeric(df_work[col], errors='coerce').fillna(0).astype(int)
    
    import unicodedata
    import re
    def get_clean_key_v4(r):
        n = unicodedata.normalize('NFC', str(r[item]))
        o = unicodedata.normalize('NFC', str(r[opt]))
        return re.sub(r'[^a-zA-Z0-9가-힣]', '', n).upper() + "_" + re.sub(r'[^a-zA-Z0-9가-힣]', '', o).upper()

    df_work['clean_key'] = df_work.apply(get_clean_key_v4, axis=1)
    
    # 데이터 매핑
    df_work["리오더 총합"] = df_work['clean_key'].map(reorder_map).fillna(0).astype(int)
    df_work["과거입고데이터"] = df_work['clean_key'].map(history_map).fillna(0).astype(int)
    df_work["리오더 입고"] = 0 
    df_work['일판매'] = df_work.apply(lambda r: int(round(r[t7]/7)) if r[t7]>0 else (int(round(r[t3]/3)) if r[t3]>0 else 0), axis=1)
    df_work['발주권장'] = ((df_work['일판매'] * (lt + ss)) - (df_work[avl] + df_work['리오더 총합'])).clip(lower=0).astype(int)

    # [4] 필터링
    is_so = df_work[s_out].astype(str).str.contains('품절', na=False)
    df_f = df_work[~is_so] if f_mode == "정상만" else (df_work[is_so] if f_mode == "품절만" else df_work)
    if s_query:
        df_f = df_f[df_f[item].astype(str).str.contains(s_query, case=False) | df_f[opt].astype(str).str.contains(s_query, case=False)]

    # [5] 화면 표시용 컬럼명 변경
    df_disp = df_f.rename(columns={
        s_out: "상태", vnd: "공급처", item: "상품명", opt: "옵션", 
        v_it: "공급처상품명", stk: "정상재고", avl: "가용재고", t3: "3일발주데이터"
    })
    
    # 4단계 테이블 최종 출력
    with st.form("v4_form"):
        target_cols = [
            "상태", "공급처", "상품명", "옵션", "공급처상품명", 
            "정상재고", "가용재고", "리오더 총합", "리오더 입고", 
            "과거입고데이터", "3일발주데이터", "일판매", "발주권장"
        ]
        
        v4_ed = st.data_editor(
            df_disp[target_cols], 
            use_container_width=True, 
            hide_index=True, 
            key="v4_editor",
            column_config={
                "정상재고": st.column_config.NumberColumn("🏢 정상재고", format="%d"),
                "가용재고": st.column_config.NumberColumn("✅ 가용재고", format="%d"),
                "리오더 총합": st.column_config.NumberColumn("📦 리오더잔량", format="%d"),
                "리오더 입고": st.column_config.NumberColumn("📥 입고차감", min_value=0),
                "과거입고데이터": st.column_config.NumberColumn("📜 과거입고", format="%d"),
                "3일발주데이터": st.column_config.NumberColumn("📅 3일판매", format="%d"),
                "일판매": st.column_config.NumberColumn("📈 일판매", format="%d"),
                "발주권장": st.column_config.NumberColumn("🚨 발주권장", format="%d")
            },
            disabled=[c for c in target_cols if c != "리오더 입고"]
        )
        
        if st.form_submit_button("💾 입고 정보 저장 및 리오더 차감", use_container_width=True):
            edits = st.session_state["v4_editor"].get("edited_rows", {})
            if edits:
                # 여기서 다시 한번 sh를 체크하여 안전하게 저장합니다.
                sh = get_sheet()
                v7_sh = sh.worksheet("발주기록")
                h_sh = sh.worksheet("입고기록")
                for r_idx, val in edits.items():
                    if "📥 입고차감" in val and int(val["📥 입고차감"]) > 0:
                        row_data = df_disp.iloc[int(r_idx)]; qty = int(val["📥 입고차감"])
                        v7_sh.append_row([datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S'), str(row_data["상품명"]), str(row_data["옵션"]), str(row_data["공급처상품명"]), 0, 0, -qty, 0, "입고차감", str(row_data["공급처"])])
                        h_sh.append_row([s_date.strftime('%Y-%m-%d'), str(row_data["상품명"]), str(row_data["옵션"]), qty])
                st.success("✅ 저장이 완료되었습니다!"); time.sleep(0.5); st.rerun()



# ------------------------------------------------------------------
# [5단계: 최종 발주 요약] - '메모' 컬럼 부재로 인한 KeyError 완벽 해결
# ------------------------------------------------------------------
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📋 5단계: 최종 발주 리스트 요약")

    # [1] 데이터 로드 및 기본 설정
    reorder_map_v5, _ = get_realtime_data_v4(datetime.now(KST).date())
    df_v5 = df_work[~df_work[s_out].astype(str).str.contains('품절', na=False)].copy()
    
    # ⭐ [핵심] df_v5와 원본 데이터 양쪽에 '메모' 컬럼이 있는지 확인하고 없으면 만듭니다.
    if '메모' not in df_v5.columns:
        df_v5['메모'] = ""
    if '메모' not in st.session_state.df_raw.columns:
        st.session_state.df_raw['메모'] = ""

    df_v5["리오더 총합"] = df_v5['clean_key'].map(reorder_map_v5).fillna(0).astype(int)
    
    if 'add_order_dict' not in st.session_state: 
        st.session_state.add_order_dict = {}
    
    df_v5['추가발주입력'] = df_v5.index.map(st.session_state.add_order_dict).fillna(0).astype(int)
    df_v5['상태표시'] = df_v5['발주권장'].apply(lambda x: "🚨 긴급" if x > 0 else "✅ 정상")

    # [2] 검색 UI
    f1, f2, f3 = st.columns([1.5, 2, 1])
    with f1: m5_f = st.selectbox("🚦 상태 필터", ["🚨 고위험/주의", "✅ 전체정상"], key="v5_f")
    with f2: s5_q = st.text_input("🔍 상품명/옵션 검색", key="v5_s")
    with f3: d5_d = st.date_input("🗓️ 발주 기록 날짜", datetime.now(KST).date(), key="v5_date")

    df_v5_v = df_v5[df_v5[item].astype(str).str.contains(s5_q, case=False)] if s5_q else df_v5
    if m5_f == "🚨 고위험/주의": 
        df_v5_v = df_v5_v[df_v5_v['발주권장'] > 0]
    
    # [3] 컬럼 매핑 (KeyError 방지를 위해 실제 존재하는 컬럼만 필터링)
    v_map = {
        "상태표시": "상태", item: "상품명", opt: "옵션", v_it: "공급처상품명",
        avl: "가용재고", "리오더 총합": "리오더잔량", "추가발주입력": "추가발주", "발주권장": "발주권장", "메모": "메모"
    }
    
    # 실제 df_v5_v에 존재하는 컬럼들만 추려서 가져옵니다.
    actual_cols = [c for c in v_map.keys() if c in df_v5_v.columns]

    with st.form("v5_form"):
        # 여기서 KeyError가 나지 않도록 actual_cols를 사용합니다.
        df_ed = df_v5_v[actual_cols].rename(columns=v_map)
        st.data_editor(
            df_ed, use_container_width=True, hide_index=True, key="v5_edit",
            column_config={
                "가용재고": st.column_config.NumberColumn("✅ 가용재고", format="%d"),
                "리오더잔량": st.column_config.NumberColumn("📦 리오더잔량", format="%d"),
                "추가발주": st.column_config.NumberColumn("➕ 추가발주", min_value=0),
                "발주권장": st.column_config.NumberColumn("🚨 발주권장", format="%d"),
                "메모": st.column_config.TextColumn("📝 메모")
            },
            disabled=[c for c in v_map.values() if c not in ["추가발주", "메모"]]
        )
        if st.form_submit_button("✅ 1. 추가발주 및 메모 확정 (저장 전 필수)", use_container_width=True):
            edits = st.session_state["v5_edit"].get("edited_rows", {})
            for r_idx, val in edits.items():
                idx = df_v5_v.index[int(r_idx)]
                if "추가발주" in val: st.session_state.add_order_dict[idx] = int(val["추가발주"])
                if "메모" in val: st.session_state.df_raw.at[idx, "메모"] = str(val["메모"])
            st.success("✅ 확정되었습니다!"); time.sleep(0.5); st.rerun()

    # [4] 저장 및 다운로드 버튼
    c_save, c_down = st.columns(2)
    with c_save:
        if st.button("💾 2. 구글 시트 최종 저장", use_container_width=True, type="primary"):
            v_ids = [k for k, v in st.session_state.add_order_dict.items() if v > 0]
            if v_ids:
                try:
                    ws_log = get_sheet().worksheet("발주기록")
                    now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
                    rows = []
                    for i in v_ids:
                        memo_val = st.session_state.df_raw.at[i, '메모'] if '메모' in st.session_state.df_raw.columns else ""
                        row = [
                            now_s, str(st.session_state.df_raw.at[i, item]), str(st.session_state.df_raw.at[i, opt]), 
                            str(st.session_state.df_raw.at[i, v_it]), int(df_v5.at[i, avl]), 
                            int(df_v5.at[i, '리오더 총합']), int(st.session_state.add_order_dict[i]), 
                            0, str(memo_val), str(st.session_state.df_raw.at[i, vnd])
                        ]
                        rows.append(row)
                    ws_log.append_rows(rows)
                    st.session_state.add_order_dict = {} 
                    st.success("✅ 구글 시트 저장 성공!"); time.sleep(1); st.rerun()
                except Exception as e:
                    st.error(f"시트 저장 실패: {e}")

    with c_down:
        v_ids_to_down = [k for k, v in st.session_state.add_order_dict.items() if v > 0]
        if v_ids_to_down:
            csv_t = df_v5[df_v5.index.isin(v_ids_to_down)].copy()
            csv_t['최종발주'] = csv_t.index.map(st.session_state.add_order_dict)
            res = csv_t[[vnd, item, opt, v_it, '최종발주']].rename(columns={vnd:"공급처", item:"상품명", opt:"옵션", v_it:"공급처상품명", "최종발주":"발주수량"})
            st.download_button("📥 3. 발주서(CSV) 다운로드", data=res.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig'), file_name=f"발주서_{d5_d.strftime('%m%d')}.csv", mime="text/csv", use_container_width=True)
        else:
            st.button("📥 3. 다운로드할 데이터 없음", disabled=True, use_container_width=True)


# ------------------------------------------------------------------
# [6단계: 전체 히스토리 관리] - 5단계 저장 구조(10개 컬럼) 완벽 반영 버전
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.subheader("📜 6단계: 전체 히스토리 관리")

    f1, f2, f3, f4 = st.columns([1.2, 0.8, 1.5, 1.5])
    
    with f1:
        today = datetime.now(KST).date()
        d_range = st.date_input("🗓️ 1. 조회 범위", value=(today, today), key="v6_date_range")
    
    with f2:
        st.write(""); st.write("") 
        search_trigger = st.button("🔍 2. 내역 조회", use_container_width=True, type="primary")

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
                    
                    # ⭐ [수정포인트] 5단계 저장 로직의 10개 컬럼 순서와 완벽 일치시킴
                    # 순서: 발주시간, 상품명, 옵션, 공급처상품명, 가용재고, 리오더잔량, 추가발주, 발주권장(0), 메모, 업체명
                    df_all.columns = [
                        "발주시간", "상품명", "옵션", "공급처상품명", 
                        "가용재고", "리오더잔량", "추가발주", "발주권장", "메모", "업체명"
                    ]
                    
                    df_all["날짜_만"] = df_all["발주시간"].astype(str).str.slice(0, 10)
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
                    
                    df_filtered = df_all[(df_all["날짜_만"] >= s_d) & (df_all["날짜_만"] <= e_d)].copy()
                    st.session_state.v6_data = df_filtered
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
        
        # 숫자 변환 (가용재고, 리오더잔량, 추가발주 포함)
        num_cols = ["가용재고", "리오더잔량", "추가발주", "발주권장"]
        for col in num_cols:
            df_display[col] = pd.to_numeric(df_display[col], errors='coerce').fillna(0)

        if sel_session_label == "📊 선택 범위 전체 합산":
            display_title = st.session_state.v6_display_text + " 발주 합계"
            # 합산 로직 (가용/권장 등은 마지막 값 기준, 추가발주는 합계)
            df_display = df_display.groupby(["업체명", "상품명", "옵션", "공급처상품명"], as_index=False).agg({
                "발주시간": "max", 
                "가용재고": "last",
                "리오더잔량": "last",
                "추가발주": "sum",
                "발주권장": "last",
                "메모": lambda x: " / ".join(set(filter(None, x.astype(str))))
            })
        else:
            target_time = st.session_state.v6_sessions[session_options.index(sel_session_label)-1]
            df_display = df_display[df_display["발주시간"] == target_time].copy()
            display_title = f"✅ {sel_session_label} 상세 내역"

        if h_q:
            df_display = df_display[
                df_display["상품명"].astype(str).str.contains(h_q, case=False) | 
                df_display["옵션"].astype(str).str.contains(h_q, case=False)
            ]

        if not df_display.empty:
            st.write(f"#### {display_title}")
            # ⭐ 화면에 보여줄 순서 (10개 컬럼 전체 노출)
            view_order = ["발주시간", "업체명", "상품명", "옵션", "공급처상품명", "가용재고", "리오더잔량", "추가발주", "발주권장", "메모"]
            st.dataframe(df_display[view_order], use_container_width=True, hide_index=True)
            
            csv_data = df_display[view_order].to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button(
                label=f"📥 {display_title} CSV 다운로드", 
                data=csv_data, 
                file_name=f"발주히스토리_{datetime.now().strftime('%m%d')}.csv", 
                mime="text/csv",
                use_container_width=True
            )

# ------------------------------------------------------------------
# [7단계: 실시간 리오더 최종 잔량 상황판] - 컬럼명 무관 강제 매핑 버전
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.subheader("🚀 7단계: 실시간 리오더 최종 잔량 상황판")

    try:
        # [1] 데이터 로드 (캐시 없이 실시간으로 시도)
        ws_v7 = get_sheet().worksheet("발주기록")
        all_v7 = ws_v7.get_all_values()
        
        if len(all_v7) > 1:
            # 데이터프레임 생성 (제목 줄 제외)
            df_raw_v7 = pd.DataFrame(all_v7[1:])
            
            # ⭐ 핵심: 컬럼명을 시트에서 가져오지 않고 코드에서 강제로 부여 (Mismatch 방지)
            # 5단계 저장 순서: 발주시간(0), 상품명(1), 옵션(2), 공급처상품명(3), 가용(4), 리오더잔량(5), 추가발주(6), 권장(7), 메모(8), 업체명(9)
            forced_cols = ["발주시간", "상품명", "옵션", "공급처상품명", "가용", "리잔량", "수량", "권장", "메모", "업체명"]
            
            # 현재 불러온 데이터의 칸 수에 맞춰서 이름을 붙여줍니다.
            df_raw_v7.columns = forced_cols[:len(df_raw_v7.columns)]
            
            # [2] 전처리: 수량(6번 칸)을 숫자로 변환
            df_raw_v7["수량"] = pd.to_numeric(df_raw_v7["수량"], errors='coerce').fillna(0).astype(int)
            df_raw_v7["업체명"] = df_raw_v7["업체명"].astype(str).str.strip()
            
            # [3] 필터 및 검색 UI
            f1, f2, f3 = st.columns([1.5, 1.5, 1])
            with f1:
                q_v7 = st.text_input("🔍 상품명 또는 옵션 검색", key="v7_search_q")
            with f2:
                v_list = sorted(df_raw_v7["업체명"].unique().tolist())
                v_choice = st.selectbox("🏭 업체 선택", ["전체 업체"] + v_list, key="v7_vendor_sel")
            with f3:
                st.write(""); st.write("")
                if st.button("🔄 데이터 새로고침", use_container_width=True):
                    st.rerun()

            # [4] 필터링 적용
            df_f = df_raw_v7[df_raw_v7["수량"] != 0].copy() # 수량이 있는 것만
            
            if v_choice != "전체 업체":
                df_f = df_f[df_f["업체명"] == v_choice]
            if q_v7:
                df_f = df_f[df_f["상품명"].str.contains(q_v7, case=False) | df_f["옵션"].str.contains(q_v7, case=False)]

            # [5] 최종 출력 (집계)
            if not df_f.empty:
                # 날짜 보기 좋게 자르기 (2026-04-01 -> 04-01)
                df_f["날짜"] = df_f["발주시간"].str.slice(5, 10)
                
                # 중복 항목 합산 및 메모 통합
                df_final = df_f.groupby(["날짜", "업체명", "상품명", "옵션", "공급처상품명"], as_index=False).agg({
                    "수량": "sum",
                    "메모": lambda x: " / ".join(set(filter(None, x.astype(str))))
                })

                # 합계 표시
                total_qty = df_final["수량"].sum()
                st.info(f"📊 현재 조건의 미입고 총 합계: **{total_qty:,}개**")

                # 테이블 출력 (너비 최적화)
                st.dataframe(
                    df_final.sort_values(["날짜", "업체명"], ascending=[False, True]),
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "날짜": st.column_config.TextColumn("📅", width=50),
                        "수량": st.column_config.NumberColumn("🔢 잔량", format="%d", width=60),
                        "메모": st.column_config.TextColumn("📝 비고", width=300)
                    }
                )
            else:
                st.warning("⚠️ 표시할 미입고 데이터가 없습니다. (수량이 0이거나 필터 조건에 없음)")
        else:
            st.error("📡 '발주기록' 시트에 데이터가 없습니다. 5단계에서 저장을 먼저 진행해 주세요.")

    except Exception as e:
        st.error(f"❌ 7단계 로드 중 치명적 오류: {e}")
