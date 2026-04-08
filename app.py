import streamlit as st
import pandas as pd
import numpy as np
import time
import io
import re  # 정규표현식 (문자열 치환용)
import unicodedata
import streamlit.components.v1 as components
import gspread
from oauth2client.service_account import ServiceAccountCredentials
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


def get_sheet():
    try:
        # secrets에서 인증 정보 로드
        creds = ServiceAccountCredentials.from_json_keyfile_info(st.secrets["gcp_service_account"], scope)
        client = gspread.authorize(creds)
        # 🚨 여기서 시트 이름이 정확해야 합니다!
        return client.open("사장님_구글_시트_이름") 
    except Exception as e:
        st.error(f"시트 연결 실패: {e}")
        return None  # 연결 실패 시 None 반환 -> 여기서 'NoneType' 에러 발생



def get_sheet():
    # 1. 권한 범위 설정
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/drive"
    ]
    
    try:
        # 2. 인증 정보 가져오기 (함수명 확인: from_json_keyfile_dict)
        # st.secrets["gcp_service_account"]가 딕셔너리 형태여야 합니다.
        creds = ServiceAccountCredentials.from_json_keyfile_dict(
            st.secrets["gcp_service_account"], 
            scope
        )
        client = gspread.authorize(creds)
        
        # 3. 구글 시트 열기 (실제 시트 이름으로 정확히 수정하세요!)
        # 예: return client.open("내발주관리시트")
        return client.open("여기에_구글_시트_이름_입력") 
        
    except Exception as e:
        # 여기서 에러 메시지를 상세히 출력해서 원인을 잡습니다.
        st.error(f"📡 시트 연결 실패 상세: {e}")
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
        

def get_realtime_data_v4(target_date):
    try:
        import unicodedata
        import re
        
        # ⭐ 모든 공백과 특수문자를 제거하여 순수 글자만 남기는 함수
        def super_clean_v4(t):
            if not t: return ""
            t = unicodedata.normalize('NFC', str(t))
            # 한글, 영문, 숫자만 남기고 싹 제거
            return re.sub(r'[^a-zA-Z0-9가-힣]', '', t).upper().strip()

        sh = get_sheet()
        
        # 1. 발주기록 매핑 (리오더 잔량)
        ws_v7 = sh.worksheet("발주기록")
        d7 = ws_v7.get_all_values()
        r_map = {}
        if len(d7) > 1:
            for row in d7[1:]:
                try:
                    # ⭐ [상품명 + 옵션] 순서로 합침
                    key = super_clean_v4(row[1]) + super_clean_v4(row[2])
                    val = int(float(str(row[5]).replace(",", ""))) if row[5] else 0
                    add = int(float(str(row[6]).replace(",", ""))) if row[6] else 0
                    if (val + add) != 0:
                        r_map[key] = r_map.get(key, 0) + (val + add)
                except: continue

        # 2. 입고기록 매핑 (과거 입고)
        ws_h = sh.worksheet("입고기록")
        dh = ws_h.get_all_values()
        h_map = {}
        t_str = target_date.strftime('%Y-%m-%d')
        
        if len(dh) > 1:
            for row_h in dh[1:]:
                try:
                    # 날짜 비교 (시간 무관하게 날짜만 포함되면 OK)
                    if t_str in str(row_h[0]):
                        # ⭐ 동일하게 [상품명 + 옵션] 합침
                        h_key = super_clean_v4(row_h[1]) + super_clean_v4(row_h[2])
                        qty = int(float(str(row_h[3]).replace(",", "")))
                        h_map[h_key] = h_map.get(h_key, 0) + qty
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
        key=f"up_file_{st.session_state.get('upload_key', 0)}"
    )
    
    # [🔄 화면 전체 초기화 버튼]
    if st.button("🔄 화면 전체 초기화", use_container_width=True):
        for key in list(st.session_state.keys()):
            if key != "upload_key": 
                del st.session_state[key]
        st.session_state.upload_key = st.session_state.get('upload_key', 0) + 1
        st.session_state.analyzed = False 
        st.session_state.df_raw = None
        st.query_params.clear() 
        st.rerun()

    # --- 데이터 로드 및 시트 동기화 로직 ---
    if up_file:
        if st.session_state.get('df_raw') is None:
            try:
                # 1. 파일 읽기
                df = pd.read_csv(up_file) if up_file.name.endswith('.csv') else pd.read_excel(up_file)
                df.columns = [str(c).strip() for c in df.columns] 
                
                # 2. 구글 시트 리오더 수량 매칭
                with st.spinner("🔄 구글 시트에서 리오더 수량을 실시간 매칭 중..."):
                    df = sync_reorder_from_sheet(df)
                
                if "리오더 수량" not in df.columns: 
                    df["리오더 수량"] = 0
                
                df = df.fillna("") 
                st.session_state.df_raw = df
                
            except Exception as e:
                st.error(f"파일 로드 오류: {e}")

    # --- 2~3단계: 데이터가 로드된 경우에만 표시 (이 안으로 3단계를 넣어야 초기화가 먹힙니다) ---
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
            
            t3 = st.selectbox("🔥 3일 판매", cols, index=auto_idx(['3일'], exclude_keys=['1주', '7일', '품절']), key="sel_t3")
            t7 = st.selectbox("📅 7일 판매", cols, index=auto_idx(['7일', '1주'], exclude_keys=['3일', '품절']), key="sel_t7")
            reg = st.selectbox("📆 상품 등록일", cols, index=auto_idx(['등록일', '등록일자', '최초등록']), key="sel_reg")

        # --- 3단계: 데이터 분석 설정 (⭐ 위치 수정: df_raw가 있을 때만 실행됨) ---
        st.divider()
        st.subheader("🚀 3단계: 데이터 분석 설정")
        s1, s2 = st.columns(2)
        with s1:
            lt_val = st.number_input("⏳ 리드타임 (일)", value=7, key="inp_lt")
        with s2:
            ss_val = st.number_input("🛡️ 안전재고 (일)", value=3, key="inp_ss")

        # 버튼 배치 (분석 시작 / 화면 초기화)
        b1, b2 = st.columns([3, 1])
        with b1:
            if st.button("📊 데이터 분석 시작", use_container_width=True, type="primary"):
                st.session_state.p = {
                    'so': so, 'vn': vn, 'vi': vi, 'it': it, 'op': op, 
                    'st': stk, 'av': av, 't3': t3, 't7': t7, 'reg': reg,
                    'lt': lt_val, 'ss': ss_val
                }
                
                df_final = st.session_state.df_raw.copy()
                if reg in df_final.columns:
                    df_final[reg] = pd.to_datetime(df_final[reg], errors='coerce')
                
                df_final["리오더 수량"] = pd.to_numeric(df_final.get("리오더 수량", 0), errors='coerce').fillna(0)

                st.session_state.df_raw = df_final 
                st.session_state.analyzed = True   
                st.rerun()
        
        with b2:
            # 🧹 [수정 핵심] 3단계 내부에서 작동하는 개별 초기화 버튼
            if st.button("🧹 화면 초기화", use_container_width=True):
                # 분석 관련 세션만 타겟팅해서 삭제
                target_keys = ['analyzed', 'p', 'df_raw', 'v6_data', 'v7_data']
                for k in target_keys:
                    if k in st.session_state:
                        del st.session_state[k]
                
                # 업로드 키 변경으로 파일 업로더 초기화
                st.session_state.upload_key = st.session_state.get('upload_key', 0) + 1
                st.rerun()


# ------------------------------------------------------------------
# [4단계: 데이터 편집 및 재고 관리] - 초과 입고 제로화 적용 버전
# ------------------------------------------------------------------
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📊 4단계: 데이터 편집 및 재고 관리")

    sh = get_sheet()
    p = st.session_state.p
    s_out, item, opt = p['so'], p['it'], p['op']
    vnd, v_it = p['vn'], p['vi']
    stk, avl, t3, t7 = p['st'], p['av'], p['t3'], p['t7']
    lt, ss = p['lt'], p['ss']

    # UI 필터
    c1, c2, c3 = st.columns([1, 2, 1])
    with c1: f_mode = st.selectbox("🚦 상태 필터", ["전체보기", "정상만", "품절만"], index=1, key="v4_filter_mode")
    with c2: s_query = st.text_input("🔍 상품 검색", key="v4_search_query")
    with c3: s_date = st.date_input("🗓️ 입고 조회 날짜", datetime.now(KST).date(), key="v4_in_date")

    # 데이터 준비
    reorder_map, history_map = get_realtime_data_v4(s_date)
    df_work = st.session_state.df_raw.copy()

    for col in [stk, avl, t3, t7]:
        df_work[col] = pd.to_numeric(df_work[col], errors='coerce').fillna(0).astype(int)

    def get_clean_key_v4(r):
        import unicodedata, re
        n = re.sub(r'[^a-zA-Z0-9가-힣]', '', unicodedata.normalize('NFC', str(r[item]))).upper().strip()
        o = re.sub(r'[^a-zA-Z0-9가-힣]', '', unicodedata.normalize('NFC', str(r[opt]))).upper().strip()
        return n + o

    df_work['clean_key'] = df_work.apply(get_clean_key_v4, axis=1)
    
    # 1. 리오더 잔량 계산
    df_work["리오더 총합"] = df_work['clean_key'].map(reorder_map).fillna(0).astype(int)
    
    # ⭐ [핵심 추가] 리오더 잔량이 마이너스인 경우 0으로 제로화 (초과 입고 처리)
    df_work["리오더 총합"] = df_work["리오더 총합"].clip(lower=0)
    
    df_work["과거입고데이터"] = df_work['clean_key'].map(history_map).fillna(0).astype(int)
    df_work["리오더 입고"] = 0 
    
    # 2. 일판매 및 발주권장 계산 (제로화된 리오더 총합 기준)
    df_work['일판매'] = df_work.apply(lambda r: int(round(r[t7]/7)) if r[t7]>0 else (int(round(r[t3]/3)) if r[t3]>0 else 0), axis=1)
    df_work['발주권장'] = ((df_work['일판매'] * (lt + ss)) - (df_work[avl] + df_work['리오더 총합'])).clip(lower=0).astype(int)

    is_so = df_work[s_out].astype(str).str.contains('품절', na=False)
    df_f = df_work[~is_so] if f_mode == "정상만" else (df_work[is_so] if f_mode == "품절만" else df_work)
    if s_query:
        df_f = df_f[df_f[item].astype(str).str.contains(s_query, case=False) | df_f[opt].astype(str).str.contains(s_query, case=False)]

    df_disp = df_f.rename(columns={s_out: "상태", vnd: "공급처", item: "상품명", opt: "옵션", v_it: "공급처상품명", stk: "정상재고", avl: "가용재고", t3: "3일발주"})
    
    # 폼 내부 에디터 설정
    with st.form("v4_storage_form", clear_on_submit=True):
        target_cols = ["상태", "공급처", "상품명", "옵션", "공급처상품명", "정상재고", "가용재고", "리오더 총합", "리오더 입고", "과거입고데이터", "3일발주", "일판매", "발주권장"]
        
        v4_ed = st.data_editor(
            df_disp[target_cols], 
            use_container_width=True, 
            hide_index=True, 
            key="v4_main_editor", 
            column_config={
                "리오더 총합": st.column_config.NumberColumn("📦 리오더잔량"),
                "리오더 입고": st.column_config.NumberColumn("📥 입고차감", min_value=0),
                "과거입고데이터": st.column_config.NumberColumn("📜 과거입고")
            }, 
            disabled=[c for c in target_cols if c != "리오더 입고"]
        )
        
        submit_btn = st.form_submit_button("💾 입고 정보 저장 및 리오더 차감", use_container_width=True)
        
        if submit_btn:
            edits = st.session_state.get("v4_main_editor", {}).get("edited_rows", {})
            
            if edits:
                try:
                    v7_sh = sh.worksheet("발주기록")
                    h_sh = sh.worksheet("입고기록")
                    t_date = s_date.strftime('%Y-%m-%d')
                    
                    saved_count = 0
                    for r_idx, val in edits.items():
                        qty = int(val.get("리오더 입고", 0))
                        if qty > 0:
                            row = df_disp.iloc[int(r_idx)]
                            
                            # 1. 발주기록 차감 (G열에 -수량, I열에 입고차감)
                            v7_sh.append_row([
                                datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S'), 
                                str(row["상품명"]), str(row["옵션"]), str(row.get("공급처상품명", "")), 
                                0, 0, -qty, 0, "입고차감", str(row.get("공급처", "미지정"))
                            ])
                            
                            # 2. 입고기록 저장
                            h_sh.append_row([t_date, str(row["상품명"]), str(row["옵션"]), qty])
                            saved_count += 1
                    
                    if saved_count > 0:
                        st.success(f"✅ {saved_count}건 저장 완료! (시트 반영됨)")
                        st.cache_data.clear() 
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.warning("입력된 수량이 0입니다.")
                except Exception as e:
                    st.error(f"❌ 저장 중 오류 발생: {e}")
            else:
                st.warning("⚠️ 수정된 내용이 없습니다. 수량을 입력한 뒤 버튼을 눌러주세요.")



# ------------------------------------------------------------------
# [5단계: 최종 발주 요약] - 중복 합산 방지 + 시트 저장 + CSV 다운로드 포함
# ------------------------------------------------------------------
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📋 5단계: 최종 발주 리스트 요약")

    # [1] 실시간 잔량 로드 (시트 G열 합계 가져오기)
    reorder_map_v5, _ = get_realtime_data_v4(datetime.now(KST).date())
    df_v5 = st.session_state.df_raw.copy()

    # 상품 식별 키 생성
    def get_clean_key_v5(r):
        import unicodedata, re
        n = re.sub(r'[^a-zA-Z0-9가-힣]', '', unicodedata.normalize('NFC', str(r[item]))).upper().strip()
        o = re.sub(r'[^a-zA-Z0-9가-힣]', '', unicodedata.normalize('NFC', str(r[opt]))).upper().strip()
        return n + o

    df_v5['clean_key'] = df_v5.apply(get_clean_key_v5, axis=1)
    for col in [stk, avl, t3, t7]:
        df_v5[col] = pd.to_numeric(df_v5[col], errors='coerce').fillna(0).astype(int)

    # 품절 제외
    df_v5 = df_v5[~df_v5[s_out].astype(str).str.contains('품절', na=False)]
    
    if '메모' not in df_v5.columns: df_v5['메모'] = ""
    if 'add_order_dict' not in st.session_state: st.session_state.add_order_dict = {}

    # [2] 화면 표시용 계산
    df_v5["기존 리오더"] = df_v5['clean_key'].map(reorder_map_v5).fillna(0).astype(int)
    df_v5['추가발주입력'] = df_v5.index.map(st.session_state.add_order_dict).fillna(0).astype(int)
    df_v5['총 리오더 합계'] = df_v5["기존 리오더"] + df_v5['추가발주입력']

    # 발주 권장 계산
    df_v5['일판매'] = df_v5.apply(lambda r: int(round(r[t7]/7)) if r[t7]>0 else (int(round(r[t3]/3)) if r[t3]>0 else 0), axis=1)
    df_v5['발주권장'] = ((df_v5['일판매'] * (lt + ss)) - (df_v5[avl] + df_v5["기존 리오더"])).clip(lower=0).astype(int)

    # [3] 필터 및 에디터
    f1, f2 = st.columns([1, 2])
    with f1: m5_f = st.selectbox("🚦 상태 필터", ["✅ 전체보기", "🚨 발주필요"], key="v5_f")
    with f2: s5_q = st.text_input("🔍 검색", key="v5_s")

    df_v5_v = df_v5.copy()
    if s5_q:
        df_v5_v = df_v5_v[df_v5_v[item].astype(str).str.contains(s5_q, case=False) | df_v5_v[opt].astype(str).str.contains(s5_q, case=False)]
    if m5_f == "🚨 발주필요": 
        df_v5_v = df_v5_v[df_v5_v['발주권장'] > 0]
    
    v_map = {
        item: "상품명", opt: "옵션", avl: "가용재고", 
        "기존 리오더": "기존잔량", "추가발주입력": "추가발주", "총 리오더 합계": "총합계", "메모": "메모"
    }
    actual_cols = [c for c in v_map.keys() if c in df_v5_v.columns]

    # --- 에디터 폼 시작 ---
    with st.form("v5_comprehensive_form"):
        df_ed = df_v5_v[actual_cols].rename(columns=v_map)
        st.data_editor(
            df_ed, use_container_width=True, hide_index=True, key="v5_editor_final",
            column_config={
                "추가발주": st.column_config.NumberColumn("➕ 추가발주", min_value=0),
                "기존잔량": st.column_config.NumberColumn("📦 기존잔량", disabled=True),
                "총합계": st.column_config.NumberColumn("📊 총합계", disabled=True),
            }
        )
        
        # 1번 버튼: 세션에 수량 확정 (화면 갱신용)
        confirm_btn = st.form_submit_button("✅ 1. 수량 확정 및 화면 갱신", use_container_width=True)
        if confirm_btn:
            edits = st.session_state.v5_editor_final.get("edited_rows", {})
            for r_idx, val in edits.items():
                idx = df_v5_v.index[int(r_idx)]
                if "추가발주" in val: st.session_state.add_order_dict[idx] = int(val["추가발주"])
                if "메모" in val: st.session_state.df_raw.at[idx, "메모"] = str(val["메모"])
            st.rerun()

    # [4] 저장 및 다운로드 (폼 외부에 배치하여 가독성 증대)
    c_save, c_down = st.columns(2)
    
    with c_save:
        if st.button("💾 2. 구글 시트 최종 저장", use_container_width=True, type="primary"):
            # 에디터에서 타이핑한 수정사항 직접 가져오기
            raw_edits = st.session_state.v5_editor_final.get("edited_rows", {})
            if raw_edits:
                try:
                    ws_log = get_sheet().worksheet("발주기록")
                    now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
                    rows_to_add = []
                    for r_idx, changes in raw_edits.items():
                        if "추가발주" in changes and changes["추가발주"] > 0:
                            real_idx = df_v5_v.index[int(r_idx)]
                            # ⭐ 핵심: 사장님이 입력한 순수 숫자만 저장 (중복 합산 방지)
                            pure_qty = int(changes["추가발주"]) 
                            
                            rows_to_add.append([
                                now_s, str(df_v5_v.at[real_idx, item]), str(df_v5_v.at[real_idx, opt]), 
                                str(df_v5_v.at[real_idx, v_it]), int(df_v5_v.at[real_idx, avl]), 
                                int(df_v5_v.at[real_idx, '기존 리오더']), pure_qty, 
                                int(df_v5_v.at[real_idx, '발주권장']), 
                                str(changes.get("메모", df_v5_v.at[real_idx, "메모"])),
                                str(df_v5_v.at[real_idx, vnd])
                            ])
                    
                    if rows_to_add:
                        ws_log.append_rows(rows_to_add)
                        st.success(f"✅ {len(rows_to_add)}건 시트 저장 완료!")
                        st.cache_data.clear()
                        st.session_state.add_order_dict = {} 
                        time.sleep(1); st.rerun()
                except Exception as e:
                    st.error(f"시트 오류: {e}")
            else:
                st.warning("먼저 '추가발주' 수량을 입력하고 [1. 수량 확정]을 눌러주세요.")

    with c_down:
        # 다운로드용 데이터 준비 (추가발주가 입력된 항목들만)
        download_list = [k for k, v in st.session_state.add_order_dict.items() if v > 0]
        if download_list:
            df_down = df_v5[df_v5.index.isin(download_list)].copy()
            df_down['최종발주수량'] = df_down.index.map(st.session_state.add_order_dict)
            
            # 발주서 양식 정리
            csv_data = df_down[[vnd, item, opt, v_it, '최종발주수량']].rename(columns={
                vnd: "공급처", item: "상품명", opt: "옵션", v_it: "공급처상품명", "최종발주수량": "수량"
            })
            
            st.download_button(
                label="📥 3. 발주서(CSV) 다운로드",
                data=csv_data.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig'),
                file_name=f"발주서_{datetime.now(KST).strftime('%m%d_%H%M')}.csv",
                mime="text/csv",
                use_container_width=True
            )
        else:
            st.button("📥 3. 다운로드 (입력값 없음)", disabled=True, use_container_width=True)



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
# [7단계: 실시간 리오더 최종 잔량 상황판] - 제로화 및 메모 보정 통합
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.subheader("🚀 7단계: 실시간 리오더 최종 잔량 상황판")

    @st.cache_data(ttl=3600)
    def get_v7_data_cached():
        try:
            ws_v7 = get_sheet().worksheet("발주기록")
            return ws_v7.get_all_values()
        except: return []

    try:
        all_v7 = get_v7_data_cached()
        if len(all_v7) > 1:
            df_v7 = pd.DataFrame(all_v7[1:])
            df_v7.columns = ["발주시간", "상품명", "옵션", "공급처상품명", "가용재고", "기존리오더", "추가발주", "발주권장", "메모", "업체명"][:len(df_v7.columns)]
            
            # 1. 수치 변환
            df_v7["기존리오더"] = pd.to_numeric(df_v7["기존리오더"], errors='coerce').fillna(0).astype(int)
            df_v7["추가발주"] = pd.to_numeric(df_v7["추가발주"], errors='coerce').fillna(0).astype(int)
            # 수치상 최종 잔량 계산
            df_v7["최종잔량"] = df_v7["기존리오더"] + df_v7["추가발주"]
            df_v7["날짜_순수"] = df_v7["발주시간"].str.slice(0, 10)

            # 2. 메모 보정: "입고차감"만 적힌 경우 수량을 붙여줌
            def fix_memo(row):
                m = str(row['메모']).strip()
                q = row['추가발주']
                if q < 0 and m == "입고차감":
                    return f"{q}개 입고차감"
                return m
            
            df_v7["메모"] = df_v7.apply(fix_memo, axis=1)

            # 필터 UI
            f1, f2, f3, f4 = st.columns([1.2, 0.6, 1.5, 1.2])
            with f1: d_range = st.date_input("🗓️ 기간", value=((datetime.now(KST)-timedelta(days=30)).date(), datetime.now(KST).date()), key="v7_dr")
            with f2:
                st.write(""); st.write("")
                if st.button("🔄 업데이트", use_container_width=True, type="primary"):
                    st.cache_data.clear()
                    st.rerun()
            with f3: q_v7 = st.text_input("🔍 검색", key="v7_qs")
            with f4: v_choice = st.selectbox("🏭 업체", ["전체 업체"] + sorted(df_v7["업체명"].unique().tolist()), key="v7_vs")

            # 필터링 로직
            df_f = df_v7.copy()
            if isinstance(d_range, (list, tuple)) and len(d_range) == 2:
                df_f = df_f[(df_f["날짜_순수"] >= d_range[0].strftime('%Y-%m-%d')) & (df_f["날짜_순수"] <= d_range[1].strftime('%Y-%m-%d'))]
            if v_choice != "전체 업체": df_f = df_f[df_f["업체명"] == v_choice]
            if q_v7: df_f = df_f[df_f["상품명"].str.contains(q_v7, case=False) | df_f["옵션"].str.contains(q_v7, case=False)]

            if not df_f.empty:
                # 3. 상세 리스트 그룹화 및 합산
                df_final = df_f.groupby(["업체명", "상품명", "옵션", "공급처상품명"], as_index=False).agg({
                    "발주시간": "max", 
                    "최종잔량": "sum",
                    "메모": lambda x: " / ".join(dict.fromkeys(filter(None, x.astype(str))))
                })

                # ⭐ [핵심 추가] 최종 합산 잔량이 마이너스이면 0으로 제로화
                df_final["최종잔량"] = df_final["최종잔량"].clip(lower=0)
                
                # 잔량이 0보다 큰 것들 위주로 정렬 (입고 완료된 건은 숨기고 싶으면 아래 주석 해제)
                # df_final = df_final[df_final["최종잔량"] > 0]
                df_final = df_final.sort_values(["발주시간", "업체명"], ascending=[False, True])

                # 4. 전광판 합산 (제로화된 수량 기준)
                st.write("### 📊 업체별 미입고 현황")
                df_v_sum = df_final.groupby("업체명")["최종잔량"].sum().reset_index().sort_values("최종잔량", ascending=False)
                df_v_sum = df_v_sum[df_v_sum["최종잔량"] > 0] # 합계가 0인 업체는 표시 제외
                
                v_cols = st.columns(4)
                for i, r in enumerate(df_v_sum.itertuples()):
                    with v_cols[i % 4]: st.metric(label=r.업체명, value=f"{int(r.최종잔량):,} 개")
                
                st.write("#### 📋 상세 리스트")
                st.dataframe(df_final, use_container_width=True, hide_index=True,
                    column_config={
                        "발주시간": st.column_config.TextColumn("🕒 최종발주"),
                        "최종잔량": st.column_config.NumberColumn("🔢 잔량", format="%d"), 
                        "메모": st.column_config.TextColumn("📝 비고(차감내역)", width=400)
                    })
            else: st.info("데이터가 없습니다.")
    except Exception as e: st.error(f"7단계 오류: {e}")
