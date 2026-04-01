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
# [4단계: 데이터 편집 및 재고 관리] - 사장님 시트 열 순서(A~J) 정밀 타격 버전
# ------------------------------------------------------------------
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📊 4단계: 데이터 편집 및 재고 관리")

    p = st.session_state.p
    s_out, item, opt = p['so'], p['it'], p['op']
    vnd, v_it = p['vn'], p['vi']
    stk, avl, t3, t7 = p['st'], p['av'], p['t3'], p['t7']
    lt, ss = p['lt'], p['ss']

    # [1] 실시간 데이터 로드 (사진 속 F, G열 합산 로직)
    def get_realtime_data_v4(target_date):
        try:
            ws_v7 = get_sheet().worksheet("발주기록")
            d7 = ws_v7.get_all_values()
            r_map = {}
            if len(d7) > 1:
                # 헤더 제외 데이터프레임 생성
                df7 = pd.DataFrame(d7[1:], columns=[c.strip() for c in d7[0]])
                
                # ⭐ 사진 기준: F열(기존리오더), G열(추가발주) 인덱스 강제 지정
                # 컬럼명이 달라도 위치로 잡아냅니다 (0부터 시작하므로 F=5, G=6)
                c_f = df7.columns[5] # 기존리오더
                c_g = df7.columns[6] # 추가발주
                
                df7[c_f] = pd.to_numeric(df7[c_f], errors='coerce').fillna(0)
                df7[c_g] = pd.to_numeric(df7[c_g], errors='coerce').fillna(0)
                df7['합계'] = df7[c_f] + df7[c_g]
                
                # 상품명+옵션 결합 (공백 제거 필수)
                df7['key'] = df7['상품명'].astype(str).str.strip() + df7['옵션'].astype(str).str.strip()
                r_map = df7.groupby('key')['합계'].sum().to_dict()

            # 입고기록 조회
            ws_h = get_sheet().worksheet("입고기록")
            dh = ws_h.get_all_values()
            h_map = {}
            if len(dh) > 1:
                dfh = pd.DataFrame(dh[1:], columns=[c.strip() for c in dh[0]])
                t_str = target_date.strftime('%Y-%m-%d')
                dfh_f = dfh[dfh['날짜'].astype(str).str.contains(t_str)]
                if not dfh_f.empty:
                    dfh_f['수량'] = pd.to_numeric(dfh_f['수량'], errors='coerce').fillna(0)
                    dfh_f['key'] = dfh_f['상품명'].astype(str).str.strip() + dfh_f['옵션'].astype(str).str.strip()
                    h_map = dfh_f.groupby('key')['수량'].sum().to_dict()
            return r_map, h_map
        except Exception as e:
            return {}, {}

    # UI 구성
    c1, c2, c3 = st.columns([1, 2, 1])
    with c1: f_mode = st.selectbox("🚦 상태 필터", ["전체보기", "정상만", "품절만"], index=1, key="v4_f")
    with c2: s_query = st.text_input("🔍 상품 검색", key="v4_s")
    with c3: s_date = st.date_input("🗓️ 입고 조회 날짜", datetime.now(KST).date(), key="v4_d")

    # 데이터 연동
    reorder_map, history_map = get_realtime_data_v4(s_date)
    df_work = st.session_state.df_raw.copy()

    # 숫자 변환
    for col in [stk, avl, t3, t7]:
        df_work[col] = pd.to_numeric(df_work[col], errors='coerce').fillna(0).astype(int)
    
    # ⭐ 리오더 잔량 매칭
    df_work['key'] = df_work[item].astype(str).str.strip() + df_work[opt].astype(str).str.strip()
    df_work["리오더 총합"] = df_work['key'].map(reorder_map).fillna(0).astype(int)
    df_work["과거 입고"] = df_work['key'].map(history_map).fillna(0).astype(int)
    df_work["리오더 입고"] = 0 
    
    # 판매량 및 권장수량 계산
    df_work['일판매'] = df_work.apply(lambda r: int(round(r[t7]/7)) if r[t7]>0 else (int(round(r[t3]/3)) if r[t3]>0 else 0), axis=1)
    df_work['3일판매'] = (df_work['일판매'] * 3).astype(int)
    df_work['권장수량'] = ((df_work['일판매'] * (lt + ss)) - (df_work[avl] + df_work['리오더 총합'])).clip(lower=0).astype(int)

    # 필터 적용
    is_so = df_work[s_out].astype(str).str.contains('품절', na=False)
    df_f = df_work[~is_so] if f_mode == "정상만" else (df_work[is_so] if f_mode == "품절만" else df_work)
    if s_query:
        df_f = df_f[df_f[item].astype(str).str.contains(s_query, case=False) | df_f[opt].astype(str).str.contains(s_query, case=False)]

    # 화면 출력
    df_disp = df_f.rename(columns={s_out:"상태", vnd:"공급쳐", v_it:"공급처상품명", item:"상품명", opt:"옵션", stk:"정상", avl:"가용"})
    final_cols = ["상태", "공급쳐", "상품명", "옵션", "공급처상품명", "정상", "가용", "리오더 총합", "리오더 입고", "과거 입고", "3일판매", "일판매", "권장수량"]

    with st.form("v4_form"):
        v4_ed = st.data_editor(df_disp[final_cols], use_container_width=True, hide_index=True, key="v4_editor",
                               column_config={"리오더 입고": st.column_config.NumberColumn("입고차감", min_value=0)})
        if st.form_submit_button("💾 입고 정보 저장 및 리오더 차감"):
            edits = st.session_state["v4_editor"].get("edited_rows", {})
            if edits:
                v7_sh, h_sh = get_sheet().worksheet("발주기록"), get_sheet().worksheet("입고기록")
                for r_idx, val in edits.items():
                    if "리오더 입고" in val and int(val["리오더 입고"]) > 0:
                        row = df_disp.iloc[int(r_idx)]
                        qty = int(val["리오더 입고"])
                        # 사진 구조(10개 컬럼)에 맞춰 -값 저장
                        v7_sh.append_row([datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S'), str(row["상품명"]), str(row["옵션"]), str(row["공급처상품명"]), 0, 0, -qty, 0, "입고차감", str(row["공급쳐"])])
                        h_sh.append_row([s_date.strftime('%Y-%m-%d'), str(row["상품명"]), str(row["옵션"]), qty])
                st.success("✅ 완료!"); time.sleep(0.5); st.rerun()



# ------------------------------------------------------------------
# [5단계: 최종 발주 요약] - 사진 속 B, C열 매칭 및 J열 업체명 저장 버전
# ------------------------------------------------------------------
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📋 5단계: 최종 발주 리스트 요약")

    # [1] 최신 리오더 잔량 다시 로드
    reorder_map_v5, _ = get_realtime_data_v4(datetime.now(KST).date())

    # [2] 데이터 준비
    df_v5 = df_work[~df_work[s_out].astype(str).str.contains('품절', na=False)].copy()
    df_v5['key'] = df_v5[item].astype(str).str.strip() + df_v5[opt].astype(str).str.strip()
    df_v5["리오더 총합"] = df_v5['key'].map(reorder_map_v5).fillna(0).astype(int)
    
    if 'add_order_dict' not in st.session_state: st.session_state.add_order_dict = {}
    df_v5['추가발주입력'] = df_v5.index.map(st.session_state.add_order_dict).fillna(0).astype(int)
    df_v5['권장수량'] = ((df_v5['일판매'] * (lt + ss)) - (df_v5[avl] + df_v5['리오더 총합'])).clip(lower=0).astype(int)
    df_v5['상태표시'] = df_v5['권장수량'].apply(lambda x: "🚨 긴급" if x > 0 else "✅ 정상")

    # 필터/검색 UI
    f1, f2, f3 = st.columns([1.5, 2, 1])
    with f1: m5_f = st.selectbox("🚦 상태 필터", ["🚨 고위험/주의", "✅ 전체정상"], key="v5_f")
    with f2: s5_q = st.text_input("🔍 상품명/옵션 검색", key="v5_s")
    with f3: d5_d = st.date_input("🗓️ 발주 기록 날짜", datetime.now(KST).date(), key="v5_date")

    df_v5_v = df_v5[df_v5[item].astype(str).str.contains(s5_q, case=False)] if s5_q else df_v5
    if m5_f == "🚨 고위험/주의":
        df_v5_v = df_v5_v[df_v5_v['권장수량'] > 0]
    
    # [3] 에디터 설정
    v_map = {"상태표시":"상태", item:"상품명", opt:"옵션", v_it:"공급처상품명", "리오더 총합":"리오더총합", "추가발주입력":"추가발주", "권장수량":"권장수량", "비고":"메모"}
    
    with st.form("v5_form"):
        df_ed = df_v5_v[list(v_map.keys())].rename(columns=v_map)
        st.data_editor(df_ed, use_container_width=True, hide_index=True, key="v5_edit",
                       column_config={"리오더총합": "기존잔량", "추가발주": "이번발주", "권장수량": "추가필요"},
                       disabled=[c for c in v_map.values() if c not in ["추가발주", "메모"]])
        
        if st.form_submit_button("✅ 1. 추가발주 및 메모 확정"):
            edits = st.session_state["v5_edit"].get("edited_rows", {})
            for r_idx, val in edits.items():
                idx = df_v5_v.index[int(r_idx)]
                if "추가발주" in val: st.session_state.add_order_dict[idx] = int(val["추가발주"])
                if "메모" in val: st.session_state.df_raw.at[idx, "비고"] = str(val["메모"])
            st.success("확정되었습니다!"); time.sleep(0.5); st.rerun()

    # [4] 저장 및 다운로드
    c_save, c_down = st.columns(2)
    with c_save:
        if st.button("💾 2. 구글 시트 최종 저장", use_container_width=True, type="primary"):
            v_ids = [k for k, v in st.session_state.add_order_dict.items() if v > 0]
            if v_ids:
                ws_log = get_sheet().worksheet("발주기록")
                now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
                # ⭐ 사진의 A~J열 순서에 100% 맞춤
                rows = [[now_s, str(st.session_state.df_raw.at[i, item]), str(st.session_state.df_raw.at[i, opt]), str(st.session_state.df_raw.at[i, v_it]), int(df_v5.at[i, avl]), int(df_v5.at[i, '리오더 총합']), int(st.session_state.add_order_dict[i]), 0, str(st.session_state.df_raw.at[i, '비고']), str(st.session_state.df_raw.at[i, vnd])] for i in v_ids]
                ws_log.append_rows(rows)
                st.session_state.add_order_dict = {}
                st.success("✅ 시트 저장 성공!"); time.sleep(1); st.rerun()

    with c_down:
        csv_t = df_v5[df_v5.index.isin(st.session_state.add_order_dict.keys())].copy()
        csv_t['최종발주'] = csv_t.index.map(st.session_state.add_order_dict)
        if not csv_t[csv_t['최종발주']>0].empty:
            res = csv_t[csv_t['최종발주']>0][[vnd, item, opt, v_it, '최종발주']].rename(columns={vnd:"공급처", item:"상품명", opt:"옵션", v_it:"공급처상품명", "최종발주":"발주수량"})
            st.download_button("📥 발주서(CSV) 다운로드", data=res.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig'), file_name=f"발주서_{d5_d.strftime('%m%d')}.csv", mime="text/csv")
            

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
# [7단계: 실시간 리오더 최종 잔량 상황판] - 컬럼 너비 최적화 (메모란 확보)
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.subheader("🚀 7단계: 실시간 리오더 최종 잔량 상황판")

    try:
        # [1. 데이터 실시간 로드]
        ws_v7 = get_sheet().worksheet("발주기록")
        all_v7 = ws_v7.get_all_values()
        
        if len(all_v7) > 1:
            header_v7 = [c.strip() for c in all_v7[0]]
            df_raw_v7 = pd.DataFrame(all_v7[1:], columns=header_v7)
            
            # 업체명 공백 제거 및 수량 숫자 변환
            if "업체명" in df_raw_v7.columns:
                df_raw_v7["업체명"] = df_raw_v7["업체명"].astype(str).str.strip()
            
            qty_col = '추가발주' if '추가발주' in df_raw_v7.columns else ('추가' if '추가' in df_raw_v7.columns else df_raw_v7.columns[5])
            time_col = '발주시간' if '발주시간' in df_raw_v7.columns else df_raw_v7.columns[0]
            
            df_raw_v7[qty_col] = pd.to_numeric(df_raw_v7[qty_col], errors='coerce').fillna(0).astype(int)
            df_raw_v7["날짜"] = df_raw_v7[time_col].astype(str).str.slice(2, 10) # "2026-" 생략해서 너비 절약 (26-04-01 형식)

            # [2. 상단 필터 영역]
            f1, f2, f3, f4 = st.columns([1.2, 0.8, 1.5, 1.5])
            with f1:
                default_start = (datetime.now(KST) - timedelta(days=60)).date() 
                d_range_v7 = st.date_input("🗓️ 1. 기간 선택", value=(default_start, datetime.now(KST).date()), key="v7_range")
            with f2:
                st.write(""); st.write("")
                if st.button("📈 2. 데이터 동기화", use_container_width=True, type="primary"):
                    st.rerun()
            with f3:
                q_v7 = st.text_input("🔍 3. 상품/옵션 검색", key="v7_search_q")
            with f4:
                v_list = sorted(df_raw_v7["업체명"].unique().tolist())
                v_choice = st.selectbox("🏭 4. 업체 선택", ["전체 업체"] + v_list, key="v7_vendor_sel")

            # --- [필터링 및 집계] ---
            if isinstance(d_range_v7, (list, tuple)) and len(d_range_v7) == 2:
                s_d, e_d = d_range_v7[0].strftime('%Y-%m-%d'), d_range_v7[1].strftime('%Y-%m-%d')
            else:
                s_d = d_range_v7[0].strftime('%Y-%m-%d') if isinstance(d_range_v7, (list, tuple)) else d_range_v7.strftime('%Y-%m-%d')
                e_d = datetime.now(KST).strftime('%Y-%m-%d')

            df_f = df_raw_v7[(df_raw_v7["날짜"].str.contains(s_d[2:]) | (df_raw_v7["날짜"] >= s_d[2:]))].copy() # 간소화된 날짜 대응
            if v_choice != "전체 업체":
                df_f = df_f[df_f["업체명"] == v_choice]
            if q_v7:
                df_f = df_f[df_f["상품명"].str.contains(q_v7, case=False) | df_f["옵션"].str.contains(q_v7, case=False)]

            if not df_f.empty:
                group_cols = ["날짜", "업체명", "상품명", "옵션", "공급처상품명"]
                df_display = df_f.groupby(group_cols, as_index=False).agg({
                    qty_col: "sum",
                    "메모": lambda x: " / ".join(set(filter(None, x.astype(str))))
                }).rename(columns={qty_col: "잔량"})

                df_display = df_display[df_display["잔량"] > 0].sort_values("날짜", ascending=False)

                # [3. 상단 업체별 요약 카드]
                df_vendor_sum = df_display.groupby("업체명")["잔량"].sum().reset_index().sort_values("잔량", ascending=False)
                st.write(f"#### 🏭 업체별 미입고 상황 (총 {df_vendor_sum['잔량'].sum():,}개)")
                v_cols = st.columns(4)
                for i, r in df_vendor_sum.reset_index(drop=True).iterrows():
                    with v_cols[i % 4]:
                        st.metric(label=r["업체명"], value=f"{r['잔량']:,}개")
                st.divider()

                # ---------------------------------------------------------
                # [4. 하단 상세 내역] - 🚨 너비 최적화 적용
                # ---------------------------------------------------------
                st.write(f"#### 📋 상세 미입고 리스트")
                display_order = ["날짜", "업체명", "상품명", "옵션", "공급처상품명", "잔량", "메모"]
                actual_cols = [c for c in display_order if c in df_display.columns]
                
                st.dataframe(
                    df_display[actual_cols], 
                    use_container_width=True, 
                    hide_index=True,
                    column_config={
                        "날짜": st.column_config.TextColumn(width=80),         # 날짜 최소화
                        "업체명": st.column_config.TextColumn(width=100),       # 업체명 최소화
                        "상품명": st.column_config.TextColumn(width=180),       
                        "옵션": st.column_config.TextColumn(width=120),         # 옵션 최소화
                        "공급처상품명": st.column_config.TextColumn(width=150), 
                        "잔량": st.column_config.NumberColumn(format="%d", width=60), # 잔량 최소화
                        "메모": st.column_config.TextColumn(width=400)          # ⭐ 메모란 최대 확보
                    }
                )
            else:
                st.info("🔎 해당 조건에 맞는 미입고 데이터가 없습니다.")

        else:
            st.info("💡 발주 데이터가 없습니다.")
            
    except Exception as e:
        st.error(f"📡 화면 최적화 오류: {e}")
