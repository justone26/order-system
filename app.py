import streamlit as st
import pandas as pd
import numpy as np
import re
import unicodedata
import gspread
from datetime import datetime, timedelta, timezone
import streamlit.components.v1 as components
import time

# 1. 앱 설정 (가장 먼저 실행되어야 함)
st.set_page_config(
    page_title="저스트원 발주 시스템", 
    layout="wide", 
    initial_sidebar_state="collapsed"
)

# 2. 시간대 설정
KST = timezone(timedelta(hours=9))

# 3. 새로고침/창닫기 방지 자바스크립트 (성능과 안정성을 위해 설정 직후 배치)
components.html(
    """
    <script>
    window.addEventListener('beforeunload', function (e) {
      e.preventDefault();
      e.returnValue = '';
    });
    </script>
    """,
    height=0,
)

# --- [공통 함수 영역] ---
def super_clean(t):
    if not t: return ""
    t = unicodedata.normalize('NFC', str(t))
    return re.sub(r'[^a-zA-Z0-9가-힣]', '', t).upper().strip()

def to_i(v):
    try:
        val = str(v).replace(",", "").strip()
        return int(float(val)) if val else 0
    except: return 0

def find_idx(cols, keys):
    for i, c in enumerate(cols):
        if any(k in str(c).upper() for k in keys): return i
    return 0

def get_sheet():
    try:
        client = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        # 사장님 시트 ID
        return client.open_by_key("1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg")
    except Exception as e:
        st.error(f"📡 시트 연결 실패: {e}")
        return None

def load_reorder_data():
    """발주시트와 입고시트를 대조하여 진짜 리오더 잔량을 계산하는 함수"""
    r_map = {}
    sh = get_sheet()
    if not sh:
        return r_map

    try:
        # 1. 발주기록 가져오기
        ws_qty = sh.worksheet("발주기록")
        qty_logs = ws_qty.get_all_values()
        
        # 2. 입고기록 가져오기
        ws_in = sh.worksheet("입고기록")
        in_logs = ws_in.get_all_values()

        # 데이터가 있을 때만 계산 시작
        if len(qty_logs) > 1:
            # 발주 데이터 정리
            df_qty = pd.DataFrame(qty_logs[1:], columns=[c.strip() for c in qty_logs[0]])
            # 키 생성 (상품명 + 옵션)
            df_qty['k'] = df_qty.apply(lambda r: super_clean(r['상품명']) + super_clean(r['옵션']), axis=1)
            # 발주 총합 (수량 컬럼명 확인 필요: '발주수량' 혹은 6번 인덱스)
            qty_series = df_qty.groupby('k')['발주수량'].apply(lambda x: x.apply(to_i).sum())

            # 입고 데이터 정리
            in_series = pd.Series()
            if len(in_logs) > 1:
                df_in = pd.DataFrame(in_logs[1:], columns=[c.strip() for c in in_logs[0]])
                df_in['k'] = df_in.apply(lambda r: super_clean(r['상품명']) + super_clean(r['옵션']), axis=1)
                in_series = df_in.groupby('k')['입고수량'].apply(lambda x: x.apply(to_i).sum())

            # 3. 잔량 계산 (발주합계 - 입고합계)
            # 입고 데이터가 없는 키는 0으로 처리
            final_reorder = qty_series.sub(in_series, fill_value=0)
            
            # 마이너스 재고는 0으로 처리하고 딕셔너리로 변환
            r_map = final_reorder.clip(lower=0).to_dict()

    except Exception as e:
        print(f"리오더 잔량 계산 중 오류 발생: {e}")
        
    return r_map
    
# --- [메인 로직 시작] ---

# ------------------------------------------------------------------
# 1️⃣단계: 파일 업로드 및 데이터 로드 (B 시스템: 상시 노출 및 완전 초기화)
# ------------------------------------------------------------------
st.header("1️⃣ 파일 업로드")

# 파일 업로더 초기화를 위한 키 관리
if 'uploader_key' not in st.session_state:
    st.session_state.uploader_key = 0

up_file = st.file_uploader(
    "엑셀 파일을 업로드하세요.", 
    type=['xlsx', 'xls'],
    key=f"file_uploader_{st.session_state.uploader_key}"
)

# [🚨 핵심 수정] 화면 전체 초기화 버튼 로직
if st.button("♻️ 현재 화면 데이터 초기화", use_container_width=True):
    # 세션 상태 전체 순회하며 삭제
    for key in list(st.session_state.keys()):
        # uploader_key만 남기고 1~6단계에 쓰이는 모든 데이터를 삭제합니다.
        if key != 'uploader_key':
            del st.session_state[key]
    
    # 파일 업로더를 강제로 비우기 위해 키값 증가
    st.session_state.uploader_key += 1
    
    st.cache_data.clear()
    st.success("✅ 1~6단계 모든 데이터와 화면이 초기화되었습니다.")
    time.sleep(0.5)
    st.rerun()

# 파일 업로드 시 로직 (UI 및 기능 유지)
if up_file is not None:
    if 'df_raw' not in st.session_state:
        try:
            temp_df = pd.read_excel(up_file)
            st.session_state.df_raw = temp_df
            st.session_state.analyzed = False # 아직 3단계 분석 전임을 표시
            
            with st.spinner("🔄 구글 시트에서 최신 기존리오더 동기화 중..."):
                sh = get_sheet()
                ws = sh.worksheet("발주기록")
                raw_records = ws.get_all_values()
                
                if len(raw_records) > 1:
                    df_rec = pd.DataFrame(raw_records[1:], columns=[h.strip() for h in raw_records[0]])
                    def c_func(t): return "".join(str(t).split()).upper()
                    
                    df_rec['key'] = df_rec['상품명'].apply(c_func) + df_rec['옵션'].apply(c_func)
                    
                    # 1. 숫자 변환
                    df_rec['기존리오더'] = pd.to_numeric(df_rec['기존리오더'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                    
                    # 2. 날짜/시간 정렬 (통장 잔액 방식 정밀화)
                    df_rec['날짜_dt'] = pd.to_datetime(df_rec['날짜'], errors='coerce', format='mixed')
                    df_rec = df_rec.sort_values('날짜_dt', ascending=True)
                    
                    # 3. 최신 잔액 추출 및 세션 저장
                    last_balance_map = df_rec.drop_duplicates('key', keep='last').set_index('key')['기존리오더'].to_dict()
                    st.session_state.reorder_ans = last_balance_map
                else:
                    st.session_state.reorder_ans = {}
                    
            st.success("✅ 엑셀 및 장부 잔량 동기화 완료!")
            st.rerun()
        except Exception as e:
            st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")


# ------------------------------------------------------------------
# 2️⃣단계: 매핑 설정 (B 시스템: 상시 노출 및 틀 유지 버전)
# ------------------------------------------------------------------
st.divider()
st.subheader("2️⃣ 매핑 설정") # 제목은 파일 업로드 여부와 상관없이 항상 노출

if 'df_raw' in st.session_state:
    # 파일이 업로드된 경우에만 매핑 UI 노출
    df_work = st.session_state.df_raw
    cols = df_work.columns.tolist()
    
    st.info("💡 엑셀 컬럼과 시스템 항목을 매핑합니다. 기본값은 자동 설정됩니다.")
    c1, c2 = st.columns(2)
    with c1:
        # key를 지정하여 세션에 자동 저장되도록 보강 (UI 및 로직 유지)
        sold_out = st.selectbox("1. 품절 여부", cols, index=find_idx(cols, ['품절']), key="sel_so")
        vendor = st.selectbox("2. 공급처(업체명)", cols, index=find_idx(cols, ['공급처', '업체']), key="sel_vn")
        item = st.selectbox("3. 상품명", cols, index=find_idx(cols, ['상품명', '품명']), key="sel_it")
        option = st.selectbox("4. 옵션", cols, index=find_idx(cols, ['옵션', '규격']), key="sel_op")
        v_item_col = st.selectbox("5. 공급처 상품명", cols, index=find_idx(cols, ['공급처상품명']), key="sel_vi")
    with c2:
        reg_date = st.selectbox("6. 등록일", cols, index=find_idx(cols, ['등록일']), key="sel_rd")
        stock = st.selectbox("7. 정상재고", cols, index=find_idx(cols, ['정상재고']), key="sel_st")
        avail = st.selectbox("8. 가용재고", cols, index=find_idx(cols, ['가용재고', '현재고']), key="sel_av")
        t3d = st.selectbox("9. 3일 발주합계", cols, index=find_idx(cols, ['3일']), key="sel_t3")
        t1w = st.selectbox("10. 7일 발주합계", cols, index=find_idx(cols, ['7일', '1주']), key="sel_t7")
else:
    # 파일이 없는 경우 안내 메시지만 출력
    st.info("파일을 업로드하면 매핑 설정 화면이 나타납니다.")


# ------------------------------------------------------------------
# 3️⃣단계: 분석 설정 및 실행 (B 시스템: 상시 노출 및 분석 로직 유지)
# ------------------------------------------------------------------
st.divider()
st.subheader("3️⃣ 분석 설정 및 실행") # 제목 상시 노출

# [수정] 기존 clt, css 2칸에서 cmq를 추가한 3칸으로 변경
clt, css, cmq = st.columns(3)
with clt: 
    lt = st.number_input("리드타임 (일)", value=7, key="input_lt")
with css: 
    ss = st.number_input("안전재고 (일 수)", value=3, key="input_ss")
with cmq:
    # ✅ 추가된 기능: 최소 발주 기준 수량 설정
    min_qty = st.number_input("최소 발주 기준", value=10, step=1, key="input_min_qty")

# 분석 실행은 파일이 있을 때만 버튼 활성화 또는 작동
if 'df_raw' in st.session_state:
    if st.button("🚀 분석 실행", type="primary", use_container_width=True):
        try:
            # 1. 구글 시트에서 실시간 최신 잔액 동기화 (기존 로직 유지)
            with st.spinner("🔄 장부에서 최신 잔액(기존리오더) 동기화 중..."):
                sh = get_sheet()
                ws = sh.worksheet("발주기록")
                raw_records = ws.get_all_values()
                
                if len(raw_records) > 1:
                    df_rec = pd.DataFrame(raw_records[1:], columns=[h.strip() for h in raw_records[0]])
                    def c_func(t): return "".join(str(t).split()).upper()
                    
                    df_rec['key'] = df_rec['상품명'].apply(c_func) + df_rec['옵션'].apply(c_func)
                    df_rec['기존리오더'] = pd.to_numeric(df_rec['기존리오더'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                    df_rec['날짜_dt'] = pd.to_datetime(df_rec['날짜'], errors='coerce', format='mixed')
                    
                    # 최신 잔액 맵 생성
                    last_balance_map = df_rec.sort_values('날짜_dt').drop_duplicates('key', keep='last').set_index('key')['기존리오더'].to_dict()
                    st.session_state.reorder_ans = last_balance_map
                else:
                    st.session_state.reorder_ans = {}

            # 2. 분석용 파라미터 맵 구성 (min_qty 추가 저장)
            p_map = {
                'so': st.session_state.get('sel_so'), 'it': st.session_state.get('sel_it'),
                'op': st.session_state.get('sel_op'), 'vn': st.session_state.get('sel_vn'),
                'vi': st.session_state.get('sel_vi'), 'av': st.session_state.get('sel_av'),
                't3': st.session_state.get('sel_t3'), 't7': st.session_state.get('sel_t7'),
                'rd': st.session_state.get('sel_rd'), 'lt': lt, 'ss': ss,
                'min_qty': min_qty # ✅ 4단계에서 필터로 쓸 값
            }
            st.session_state.p = p_map

            # 3. 데이터 가공 및 수량 계산 (기존 로직 유지)
            df = st.session_state.df_raw.copy()
            def c_func_final(t): return "".join(str(t).split()).upper()
            df['key'] = df[p_map['it']].apply(c_func_final) + df[p_map['op']].apply(c_func_final)

            ans_map = st.session_state.get('reorder_ans', {})
            df['기존리오더'] = df['key'].map(ans_map).fillna(0).astype(int)

            # 일판매량 및 권장발주수량 계산
            c_av, c_t7, c_rd = p_map['av'], p_map['t7'], p_map['rd']
            df[c_av] = pd.to_numeric(df[c_av], errors='coerce').fillna(0).astype(int)
            
            today = datetime.now(KST).date()
            def calc_daily(row):
                try:
                    diff = (today - pd.to_datetime(row[c_rd]).date()).days
                    days = max(1, min(diff, 7))
                    return int(round(pd.to_numeric(row[c_t7], errors='coerce') / days, 0))
                except: return int(round(pd.to_numeric(row[c_t7], errors='coerce') / 7, 0))

            df['일판매량'] = df.apply(calc_daily, axis=1).fillna(0).astype(int)
            df['권장발주수량'] = ((df['일판매량'] * (lt + ss)) - (df[c_av] + df['기존리오더'])).clip(lower=0).astype(int)
            
            # 상태 판별 및 UI용 컬럼 추가
            df['상태'] = df.apply(lambda r: "🚫 품절" if "품절" in str(r[p_map['so']]) else ("🚨 발주필요" if r['권장발주수량'] > 0 else "✅ 정상"), axis=1)
            df['입고차감'] = 0 ; df['추가발주'] = 0 ; df['비고(처리내역)'] = ""

            if 'key' in df.columns: df = df.drop(columns=['key'])
            
            # 결과 저장 및 상태 변경
            st.session_state.df_final = df
            st.session_state.analyzed = True
            
            st.cache_data.clear()
            st.success(f"✅ 분석 완료 (최신 장고 반영 완료 / 기준: {min_qty}개 미만 제외)")
            st.rerun()

        except Exception as e:
            st.error(f"⚠️ 분석 오류: {e}")
else:
    # 파일이 없는 상태에서 분석 버튼을 누르려 할 때
    st.button("🚀 분석 실행", type="primary", use_container_width=True, disabled=True)
    st.info("1단계에서 파일을 업로드해야 분석이 가능합니다.")


# ------------------------------------------------------------------
# 4️⃣단계: 입고 관리 및 최종 저장 (필터 레이아웃 및 총 가용재고 표기)
# ------------------------------------------------------------------
st.divider()
st.header("📊 4단계: 데이터 분석 및 발주 체크") 

if st.session_state.get('analyzed'):
    p = st.session_state.p
    
    if 'df_final' not in st.session_state:
        st.error("데이터가 없습니다. 분석을 다시 진행해 주세요.")
    else:
        df_all = st.session_state.df_final.copy()

        # 컬럼 보정 (기존 로직 유지)
        for col in ['기존리오더', '입고차감', '추가발주', '일판매량', '권장발주수량', '비고(처리내역)', '상태']:
            if col not in df_all.columns:
                df_all[col] = 0 if any(x in col for x in ['수량', '리오더', '차감', '발주']) else ""

        # --- [UI 레이아웃] ---
        # 기존 3분할에서 가용재고 칸을 추가하여 4분할로 변경
        f1, f_vn, f_stock, f2 = st.columns([1, 1.2, 0.8, 1.8]) 
        
        with f1: 
            f_mode = st.selectbox("🚦 상태 필터", ["전체보기", "🚨 발주필요(세트)", "✅ 정상", "🚫 품절"], index=1, key="f_mode_4")
        
        with f_vn:
            vn_list = ["전체 공급처"] + sorted(df_all[p['vn']].unique().tolist())
            sel_vn = st.selectbox("🏭 공급처 필터", vn_list, key="f_vendor_4")
            
        with f2: 
            s_query = st.text_input("🔍 검색 (상품명/옵션)", key="s_query_4")

        # --- [데이터 필터링] ---
        df_temp = df_all.copy()
        
        # 3단계에서 정한 최소 발주 수량 미만 필터링
        min_filter = st.session_state.p.get('min_qty', 0)
        df_temp.loc[df_temp['권장발주수량'] < min_filter, '권장발주수량'] = 0
        
        # 상태 필터
        if f_mode == "🚨 발주필요(세트)":
            need_items = df_temp[(df_temp['상태'] != "🚫 품절") & (df_temp['권장발주수량'] > 0)][p['it']].unique()
            df_temp = df_temp[df_temp[p['it']].isin(need_items)]
        elif f_mode != "전체보기":
            df_temp = df_temp[df_temp['상태'] == f_mode]
        
        # 공급처 필터
        if sel_vn != "전체 공급처":
            df_temp = df_temp[df_temp[p['vn']] == sel_vn]
            
        # 검색 필터
        if s_query:
            df_temp = df_temp[df_temp[p['it']].str.contains(s_query, case=False, na=False) | 
                               df_temp[p['op']].str.contains(s_query, case=False, na=False)]

        # --- [가용재고 메트릭 표시] ---
        with f_stock:
            # 음수는 0으로 치환하여 합계 계산
            total_stock = int(df_temp[p['av']].clip(lower=0).sum())
            st.metric("총 가용재고", f"{total_stock:,} 개")

        # --- [데이터 편집기: 너비 조정 적용] ---
        full_cols = ['상태', p['vn'], p['it'], p['op'], p['vi'], p['av'], 
                     '기존리오더', '입고차감', '추가발주', p['t3'], 
                     '일판매량', '권장발주수량', '비고(처리내역)']
        disp_cols = [c for c in full_cols if c in df_temp.columns]
        
        with st.form("final_form"):
            edited_df = st.data_editor(
                df_temp[disp_cols], 
                use_container_width=True, 
                hide_index=True,
                key="main_editor", 
                column_config={
                    "상태": st.column_config.TextColumn("상태", width=70, disabled=True),
                    p['vn']: st.column_config.TextColumn("공급처", width=140, disabled=True),
                    p['it']: st.column_config.TextColumn("상품명", width=330, disabled=True),
                    p['op']: st.column_config.TextColumn("옵션", width=130, disabled=True),
                    p['vi']: st.column_config.TextColumn("공급처상품명", width=250, disabled=True),
                    "기존리오더": st.column_config.NumberColumn("기존리오더", width=80, disabled=True, format="%d"),
                    "입고차감": st.column_config.NumberColumn("📥 입고(-)", width=100, min_value=0), 
                    "추가발주": st.column_config.NumberColumn("➕ 발주(+)", width=100, min_value=0),
                    "권장발주수량": st.column_config.NumberColumn("권장수량", width=80, disabled=True, format="%d"),
                    "비고(처리내역)": st.column_config.TextColumn("비고(처리내역)", width=120),
                }
            )
            btn_save = st.form_submit_button("🚀 최종 데이터 저장 및 시트 전송", use_container_width=True, type="primary")

        # --- [저장 및 전송 로직] ---
        if btn_save:
            changed_rows = edited_df[(edited_df['입고차감'] > 0) | (edited_df['추가발주'] > 0) | (edited_df['비고(처리내역)'].str.strip() != "")].copy()
            
            if not changed_rows.empty:
                with st.spinner("🚀 장부 업데이트 중..."):
                    try:
                        sh = get_sheet()
                        ws_qty = sh.worksheet("발주기록")
                        ws_hist = sh.worksheet("히스토리")
                        now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
                        time_short = datetime.now(KST).strftime('%m/%d')
                        
                        rows_qty, rows_hist = [], []
                        for _, r in changed_rows.iterrows():
                            q_val = int(r['추가발주'])
                            i_val = int(r['입고차감'])
                            user_memo = str(r['비고(처리내역)']).strip() if r['비고(처리내역)'] and str(r['비고(처리내역)']) != "None" else ""
                            
                            bal_after_in = max(0, int(r['기존리오더']) - i_val)
                            final_balance = bal_after_in + q_val
                            
                            m_parts = []
                            if q_val > 0: m_parts.append(f"{time_short} {q_val}발주")
                            if i_val > 0: m_parts.append(f"{time_short} {i_val}입고")
                            
                            over_in_msg = f" (오버입고 {i_val - int(r['기존리오더'])}개 발생)" if i_val > int(r['기존리오더']) else ""
                            final_memo = f"[{' '.join(m_parts)}]{over_in_msg} {user_memo}".strip()
                            
                            rows_qty.append([now_s, r[p['vn']], r[p['it']], r[p['op']], r[p['vi']], final_balance, q_val, i_val, final_memo])
                            rows_hist.append([now_s, r[p['vn']], r[p['it']], r[p['op']], r[p['vi']], r[p['av']], r['기존리오더'], i_val, q_val, r['권장발주수량'], final_memo])

                            mask = (st.session_state.df_final[p['it']] == r[p['it']]) & (st.session_state.df_final[p['op']] == r[p['op']])
                            st.session_state.df_final.loc[mask, '기존리오더'] = final_balance
                            st.session_state.df_final.loc[mask, '입고차감'] = 0
                            st.session_state.df_final.loc[mask, '추가발주'] = 0

                        if rows_qty: ws_qty.append_rows(rows_qty, value_input_option='USER_ENTERED')
                        if rows_hist: ws_hist.append_rows(rows_hist, value_input_option='USER_ENTERED')

                        st.cache_data.clear()
                        st.success(f"✅ 저장 및 모든 현황판 자동 동기화 완료!")
                        time.sleep(1)
                        st.rerun() 
                    except Exception as e:
                        st.error(f"저장 중 오류 발생: {e}")
            else:
                st.warning("⚠️ 입력된 변경 내역(입고/발주/메모)이 없습니다.")
else:
    st.info("3단계에서 [분석 실행] 버튼을 누르면 분석 결과가 이곳에 표시됩니다.")



# ------------------------------------------------------------------
# 6️⃣단계: 리오더 현황판 (사장님 원본 레이아웃 + 현황판 + 간격 최적화)
# ------------------------------------------------------------------
def render_step6():
    # 상단 제목
    st.markdown("### 📈 6단계: 실시간 리오더 현황판 (상품별 통합)")
    
    # [1] 상단 컨트롤바
    c_btn, c_search, c_filter = st.columns([1, 2, 1])
    
    with c_btn:
        st.write("🔄 데이터 갱신")
        btn_update = st.button("최신 자료 업데이트", use_container_width=True, key="btn_update_final")
        if btn_update:
            st.session_state.df_log_6 = None 
            st.cache_data.clear()
            st.rerun()

    with c_search:
        st.write("🔍 통합 상품명 검색")
        sel_s = st.text_input("상품명을 입력하세요", label_visibility="collapsed", key="s6_search_final")

    # 데이터 로드 로직
    if "df_log_6" not in st.session_state or st.session_state.df_log_6 is None:
        with c_filter:
            st.write("🏭 공급처 필터")
            st.selectbox("전체 공급처", ["전체 공급처"], label_visibility="collapsed", disabled=True)
        
        try:
            with st.spinner("⏳ 데이터를 가져오는 중..."):
                sh = get_sheet()
                ws_qty = sh.worksheet("발주기록")
                data = ws_qty.get_all_values()
                if len(data) > 1:
                    headers = [h.strip() for h in data[0]]
                    df_log = pd.DataFrame(data[1:], columns=headers)
                    df_log = df_log.loc[:, ~df_log.columns.duplicated()]
                    
                    if '공급처명' in df_log.columns:
                        df_log.rename(columns={'공급처명': '공급처상품명'}, inplace=True)
                    
                    for col in ['기존리오더', '추가발주', '입고수량']:
                        if col in df_log.columns:
                            df_log[col] = pd.to_numeric(df_log[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                    
                    # 🚨 [수정] 날짜에서 초 단위 삭제 처리
                    d_col = next((c for c in df_log.columns if '날짜' in c or '시간' in c), df_log.columns[0])
                    df_log['날짜_dt'] = pd.to_datetime(df_log[d_col], errors='coerce', format='mixed')
                    
                    # 화면 표시용 날짜 포맷팅 (초 단위 제거)
                    df_log[d_col] = df_log['날짜_dt'].dt.strftime('%Y-%m-%d %H:%M')
                    
                    st.session_state.df_log_6 = df_log
                    st.rerun()
        except Exception as e:
            st.error(f"오류: {e}")
        return

    df_log = st.session_state.df_log_6.copy()
    
    with c_filter:
        st.write("🏭 공급처 필터")
        v_col = '공급처' if '공급처' in df_log.columns else df_log.columns[1]
        v_list = ["전체 공급처"] + sorted(df_log[v_col].unique().tolist())
        sel_v = st.selectbox("전체 공급처", v_list, label_visibility="collapsed", key="s6_vendor_final")

    # [2] 데이터 가공
    def c_func(t): return "".join(str(t).split()).upper()
    col_it = '상품명' if '상품명' in df_log.columns else df_log.columns[2]
    col_op = '옵션' if '옵션' in df_log.columns else df_log.columns[3]
    col_memo = '메모' if '메모' in df_log.columns else df_log.columns[-1]

    df_log['key'] = df_log[col_it].apply(c_func) + df_log[col_op].apply(c_func)
    grouped = df_log.sort_values('날짜_dt').drop_duplicates('key', keep='last').copy()
    
    memo_map = df_log.groupby('key')[col_memo].apply(
        lambda x: " / ".join([str(i).strip() for i in x.tail(5) if str(i).strip() not in ["", "None", "nan"]])
    ).to_dict()
    
    grouped['최종잔량'] = grouped['기존리오더']
    grouped['최종메모'] = grouped['key'].map(memo_map)

    filtered = grouped[grouped['최종잔량'] > 0].copy()
    if sel_s: filtered = filtered[filtered[col_it].str.contains(sel_s, case=False, na=False)]
    if sel_v != "전체 공급처": filtered = filtered[filtered[v_col] == sel_v]

    # [3] 업체별 현황판
    st.markdown("#### 🏢 업체별 미입고 및 주요 상품")
    v_sum = filtered.groupby(v_col)['최종잔량'].sum().reset_index().sort_values('최종잔량', ascending=False)
    
    if not v_sum.empty:
        v_rows = v_sum.iloc[:3]
        v_cols = st.columns(3)
        for i, (idx, row) in enumerate(v_rows.iterrows()):
            v_name = row[v_col]
            with v_cols[i]:
                st.markdown(f"**{v_name}**")
                st.markdown(f"### {int(row['최종잔량'])}개 잔량")
                v_items = filtered[filtered[v_col] == v_name]
                v_top = v_items.groupby(col_it)['최종잔량'].sum().sort_values(ascending=False).head(3)
                for rank, (s_name, s_qty) in enumerate(v_top.items()):
                    st.write(f"{rank+1}. {s_name} **({int(s_qty)})**")

    st.divider()

    # [4] 상세 데이터 표 (초 단위 없이 출력)
    display_df = filtered.sort_values(by=['날짜_dt', '최종잔량'], ascending=[False, False])
    
    d_col_name = next((c for c in display_df.columns if '날짜' in c or '시간' in c), '날짜')
    target_display = [d_col_name, '공급처', '상품명', '옵션', '공급처상품명', '최종잔량', '추가발주', '입고수량', '최종메모']
    final_cols = [c for c in target_display if c in display_df.columns]

    st.dataframe(
        display_df[final_cols].rename(columns={'최종메모': '최근 처리내역(메모)'}), 
        use_container_width=True, 
        hide_index=True,
        column_config={
            d_col_name: st.column_config.TextColumn("발주시간", width=100),
            "공급처": st.column_config.TextColumn("공급처", width=90),
            "상품명": st.column_config.TextColumn("상품명", width=350),
            "옵션": st.column_config.TextColumn("옵션", width=110),
            "공급처상품명": st.column_config.TextColumn("공급처상품명", width=250),
            "최종잔량": st.column_config.NumberColumn("최종잔량", width=60, format="%d"),
            "최근 처리내역(메모)": st.column_config.TextColumn("최근 처리내역(메모)", width=500), 
        }
    )

    # [5] 하단 엑셀 버튼
    import io
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        display_df[final_cols].to_excel(writer, index=False, sheet_name='리오더현황')
    st.download_button(label="📥 실시간 현황 엑셀 다운로드", data=output.getvalue(), 
                       file_name=f"리오더현황_{datetime.now(KST).strftime('%m%d_%H%M')}.xlsx", 
                       use_container_width=True)
    


# ------------------------------------------------------------------
# 5️⃣단계: 전체 히스토리 기록 (초 단위 완전 삭제 및 한글 요일 보강)
# ------------------------------------------------------------------
st.divider()
st.header("📜 5단계: 전체 히스토리 기록")

c1, c2, c3 = st.columns([1.5, 1.5, 1.2]) 
with c1: 
    now_dt = datetime.now(KST)
    today_val = now_dt.date()
    
    # 요일 한글 변환용 리스트
    weekday_ko = ['월', '화', '수', '목', '금', '토', '일']
    today_w = weekday_ko[today_val.weekday()]
    
    sel_dates_5 = st.date_input(
        f"📅 조회 날짜 (오늘: {today_val} {today_w}요일)", 
        [today_val, today_val], 
        key="h_date_vSplit_final",
        format="YYYY-MM-DD"
    )

# 🚨 [날짜 처리 및 한글 요일 안내 로직]
if isinstance(sel_dates_5, (list, tuple)):
    if len(sel_dates_5) == 2:
        start_date, end_date = sel_dates_5
        s_w = weekday_ko[start_date.weekday()]
        e_w = weekday_ko[end_date.weekday()]
        st.info(f"🔎 조회 범위: **{start_date}({s_w})** ~ **{end_date}({e_w})**")
    elif len(sel_dates_5) == 1:
        start_date = end_date = sel_dates_5[0]
        s_w = weekday_ko[start_date.weekday()]
        st.warning(f"📅 **{start_date}({s_w})** 선택됨 (종료일도 클릭하세요)")
    else:
        start_date = end_date = today_val
else:
    start_date = end_date = today_val

with c2: 
    h_name_5 = st.text_input("🔍 상품명/옵션 검색", key="h_name_vSplit_final")

time_select_place = c3.empty() 

# --- [데이터 로드 및 출력 로직] ---
if st.button("🔍 히스토리 데이터 불러오기", use_container_width=True, type="secondary"):
    try:
        with st.spinner("⏳ 구글 시트에서 히스토리 기록을 가져오는 중..."):
            sh = get_sheet()
            ws_hist = sh.worksheet("히스토리")
            raw_data = ws_hist.get_all_values()
            
            if len(raw_data) > 1:
                cols_5 = [c.strip() for c in raw_data[0]]
                h_df = pd.DataFrame(raw_data[1:], columns=cols_5)
                h_df = h_df.loc[:, ~h_df.columns.duplicated()]
                
                # 명칭 통일
                if '공급처명' in h_df.columns:
                    h_df.rename(columns={'공급처명': '공급처상품명'}, inplace=True)
                if '메모' in h_df.columns: h_df.rename(columns={'메모': '비고(처리내역)'}, inplace=True)
                elif '비고' in h_df.columns: h_df.rename(columns={'비고': '비고(처리내역)'}, inplace=True)
                
                # 🚨 [핵심 수정] 데이터 로드 시점에 초 단위 즉시 삭제
                d_col = next((c for c in h_df.columns if '날짜' in c or '시간' in c), h_df.columns[0])
                h_df['날짜_dt'] = pd.to_datetime(h_df[d_col], errors='coerce', format='mixed')
                # 표시용 컬럼 자체를 초 없는 포맷으로 덮어씌움
                h_df[d_col] = h_df['날짜_dt'].dt.strftime('%Y-%m-%d %H:%M')
                
                st.session_state.db_history = h_df
                st.rerun()
    except Exception as e:
        st.error(f"히스토리 로드 실패: {e}")

if "db_history" in st.session_state and not st.session_state.db_history.empty:
    m_df_5 = st.session_state.db_history.copy()
    d_col = next((c for c in m_df_5.columns if '날짜' in c or '시간' in c), m_df_5.columns[0])
    
    # 필터링용 날짜 객체 생성
    m_df_5['날짜_only'] = pd.to_datetime(m_df_5['날짜_dt']).dt.date
    
    # 1. 날짜 범위 필터링
    period_df = m_df_5[(m_df_5['날짜_only'] >= start_date) & (m_df_5['날짜_only'] <= end_date)]
    
    # 2. 회차 선택 (이미 초가 제거된 d_col 사용)
    t_opts = ["전체 회차"] + sorted(period_df[d_col].dropna().unique(), reverse=True)
    h_time_5 = time_select_place.selectbox("⏰ 저장 회차 선택", t_opts, key="h_time_vSplit_final")

    df_dis = period_df.copy()
    
    # 3. 검색어 필터링
    if h_name_5: 
        df_dis = df_dis[df_dis.apply(lambda r: h_name_5.lower() in str(r).lower(), axis=1)]
    
    # 4. 회차 필터링
    if h_time_5 != "전체 회차": 
        df_dis = df_dis[df_dis[d_col] == h_time_5]

    # 불필요한 필터용 컬럼 삭제
    final_dis_df = df_dis.sort_values(by='날짜_dt', ascending=False).drop(columns=['날짜_dt', '날짜_only'], errors='ignore')

    st.dataframe(
        final_dis_df, 
        use_container_width=True, 
        hide_index=True,
        column_config={
            "발주시간": st.column_config.TextColumn("발주시간", width=100),
            "공급처": st.column_config.TextColumn("공급처", width=90),
            "상품명": st.column_config.TextColumn("상품명", width=350),
            "옵션": st.column_config.TextColumn("옵션", width=110),
            "공급처상품명": st.column_config.TextColumn("공급처상품명", width=250),
            "최종잔량": st.column_config.NumberColumn("최종잔량", width=60),
            "비고(처리내역)": st.column_config.TextColumn("비고(처리내역)", width=400), 
            "권장수량": st.column_config.NumberColumn("권장수량", width=60),
        }
    )

    import io
    output_h = io.BytesIO()
    with pd.ExcelWriter(output_h, engine='xlsxwriter') as writer:
        final_dis_df.to_excel(writer, index=False, sheet_name='히스토리기록')
    st.download_button(label="📥 히스토리 엑셀 다운로드", data=output_h.getvalue(), 
                       file_name=f"히스토리_{datetime.now(KST).strftime('%m%d_%H%M')}.xlsx", 
                       use_container_width=True, key="btn_download_h5")
else:
    time_select_place.selectbox("⏰ 저장 회차 선택", ["전체 회차"], key="h_time_vSplit_final", disabled=True)

render_step6()
