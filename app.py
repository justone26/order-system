import streamlit as st
import pandas as pd
import numpy as np
import re
import unicodedata
import gspread
from datetime import datetime, timedelta, timezone
import streamlit.components.v1 as components
import time

# 1. 환경 및 시간 설정
KST = timezone(timedelta(hours=9))
st.set_page_config(layout="wide", page_title="저스트원 발주 시스템")

# --- [🚨 새로고침/창닫기 방지 자바스크립트] ---
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
# 1️⃣단계: 파일 업로드 및 데이터 로드
# ------------------------------------------------------------------
st.header("1️⃣ 파일 업로드 및 데이터 로드")
up_file = st.file_uploader("엑셀 파일을 업로드하세요.", type=['xlsx', 'xls'])

# ✅ 초기화 버튼: 모든 기록을 지우고 초기 상태로 만듦
if st.button("🔄 현재 화면 데이터 초기화", use_container_width=True):
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.info("✅ 초기화되었습니다. 업로드된 파일을 X 눌렀다 다시 올리거나, 새 파일을 선택해주세요.")
    st.rerun()

# 💡 수정된 로드 로직: 
# 1. 파일이 업로드되었고 
# 2. 아직 세션에 df_raw가 없을 때만 로드를 시도합니다.
if up_file is not None:
    if 'df_raw' not in st.session_state:
        try:
            # 엑셀 읽기 실행
            temp_df = pd.read_excel(up_file)
            
            # 읽기 성공 시 세션에 저장
            st.session_state.df_raw = temp_df
            st.session_state.analyzed = False
            
            # 파일이 로드되면 화면을 새로고침하여 즉시 2단계를 노출
            st.rerun()
        except Exception as e:
            st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")

# ------------------------------------------------------------------
# 2️⃣단계 & 3️⃣단계: 데이터가 로드된 상태(df_raw가 세션에 있을 때)에서만 노출
# ------------------------------------------------------------------
if 'df_raw' in st.session_state:
    st.divider()
    df_work = st.session_state.df_raw
    cols = df_work.columns.tolist()
    
    st.subheader("⚙️ 2️⃣단계: 매핑 설정")
    c1, c2 = st.columns(2)
    with c1:
        sold_out = st.selectbox("1. 품절 여부", cols, index=find_idx(cols, ['품절']))
        vendor = st.selectbox("2. 공급처(업체명)", cols, index=find_idx(cols, ['공급처', '업체']))
        item = st.selectbox("3. 상품명", cols, index=find_idx(cols, ['상품명', '품명']))
        option = st.selectbox("4. 옵션", cols, index=find_idx(cols, ['옵션', '규격']))
        v_item_col = st.selectbox("5. 공급처 상품명", cols, index=find_idx(cols, ['공급처상품명']))
    with c2:
        reg_date = st.selectbox("6. 등록일", cols, index=find_idx(cols, ['등록일']))
        stock = st.selectbox("7. 정상재고", cols, index=find_idx(cols, ['정상재고']))
        avail = st.selectbox("8. 가용재고", cols, index=find_idx(cols, ['가용재고', '현재고']))
        t3d = st.selectbox("9. 3일 발주합계", cols, index=find_idx(cols, ['3일']))
        t1w = st.selectbox("10. 7일 발주합계", cols, index=find_idx(cols, ['7일', '1주']))

    # 3️⃣단계: 분석 설정 및 실행
    st.divider()
    st.subheader("⚙️ 3️⃣단계: 분석 설정 및 실행")
    
    clt, css = st.columns(2)
    with clt: lt = st.number_input("리드타임 (일)", value=10)
    with css: ss = st.number_input("안전재고 (일 수)", value=7)

    if st.button("🚀 분석 실행 / 실시간 장부 업데이트", type="primary", use_container_width=True):
        st.session_state.p = {
            'so': sold_out, 'it': item, 'op': option, 'vn': vendor, 'vi': v_item_col,
            'av': avail, 't3': t3d, 't7': t1w, 'lt': lt, 'ss': ss, 'rd': reg_date
        }

        with st.spinner("📊 발주기록 시트 분석 및 잔량 계산 중..."):
            try:
                # 분석 실행 시 5단계 히스토리 세션 초기화 (최신 데이터 갱신을 위해)
                if 'db_history' in st.session_state:
                    del st.session_state.db_history

                df = st.session_state.df_raw.copy()
                today = datetime.now(KST).date()
                sh = get_sheet()
                
                # 시트 데이터 정리 함수
                def get_clean_df(name):
                    ws = sh.worksheet(name)
                    data = ws.get_all_values()
                    if len(data) > 1:
                        res = pd.DataFrame(data[1:], columns=[c.strip() for c in data[0]])
                        return res.loc[:, ~res.columns.duplicated()]
                    return pd.DataFrame()

                # 발주기록 로드 및 잔량 계산
                df_master = get_clean_df("발주기록")
                st.session_state.master_log = df_master 

                r_map = {}
                if not df_master.empty:
                    it_c, op_c, q_c, in_c = '상품명', '옵션', '추가발주', '입고수량'
                    if it_c in df_master.columns and op_c in df_master.columns:
                        df_master[q_c] = pd.to_numeric(df_master[q_c], errors='coerce').fillna(0)
                        df_master[in_c] = pd.to_numeric(df_master[in_c], errors='coerce').fillna(0)
                        qty_sum = df_master.groupby([it_c, op_c])[q_c].sum()
                        in_sum = df_master.groupby([it_c, op_c])[in_c].sum()
                        final_res = qty_sum.sub(in_sum, fill_value=0).clip(lower=0)
                        r_map = final_res.to_dict()

                # 분석 계산 로직
                df[avail] = pd.to_numeric(df[avail], errors='coerce').fillna(0).astype(int)
                df[t1w] = pd.to_numeric(df[t1w], errors='coerce').fillna(0).astype(int)
                
                def get_reorder_val(row):
                    k = (str(row[item]).strip(), str(row[option]).strip())
                    return int(r_map.get(k, 0))
                
                df['기존리오더'] = df.apply(get_reorder_val, axis=1)

                def get_daily_avg(row):
                    try:
                        r_dt = pd.to_datetime(row[reg_date]).date()
                        days = max(1, min((today - r_dt).days, 7))
                        return int(round(pd.to_numeric(row[t1w]) / days, 0))
                    except: return int(round(pd.to_numeric(row[t1w]) / 7, 0))

                df['일판매량'] = df.apply(get_daily_avg, axis=1)
                df['권장발주수량'] = ((df['일판매량'] * (lt + ss)) - (df[avail] + df['기존리오더'])).clip(lower=0).astype(int)
                df['상태'] = df.apply(lambda r: "🚫 품절" if "품절" in str(r[sold_out]) else ("🚨 발주필요" if r['권장발주수량'] > 0 else "✅ 정상"), axis=1)
                
                df['입고차감'] = 0  
                df['추가발주'] = 0
                df['비고(처리내역)'] = "" 
                
                st.session_state.df_final = df
                st.session_state.analyzed = True
                st.success("✅ 분석 완료!")
                st.rerun()
                
            except Exception as e:
                st.error(f"⚠️ 분석 오류: {e}")
                
                
# ------------------------------------------------------------------
# 4️⃣단계: 입고 관리 및 최종 저장 (상품별 묶음 저장 + 순서 교정)
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.header("📊 4단계: 입고 관리 및 최종 발주 확정")
    
    p = st.session_state.p
    
    # 세션 상태 유지를 위해 데이터 복사
    if 'edited_df_state' not in st.session_state:
        st.session_state.edited_df_state = st.session_state.df_final.copy()

    f1, f2 = st.columns([1, 2])
    with f1: f_mode = st.selectbox("🚦 상태 필터", ["전체보기", "🚨 발주필요(세트)", "✅ 정상", "🚫 품절"], index=1)
    with f2: s_query = st.text_input("🔍 검색 (상품명/옵션)")

    # 필터링용 데이터프레임 구성
    df_temp = st.session_state.df_final.copy()
    if f_mode == "🚨 발주필요(세트)":
        need_items = df_temp[(df_temp['상태'] != "🚫 품절") & (df_temp['권장발주수량'] > 0)][p['it']].unique()
        df_temp = df_temp[df_temp[p['it']].isin(need_items)]
    elif f_mode != "전체보기":
        df_temp = df_temp[df_temp['상태'] == f_mode]
    if s_query:
        df_temp = df_temp[df_temp[p['it']].str.contains(s_query, case=False) | df_temp[p['op']].str.contains(s_query, case=False)]

    # 화면에 보여줄 순서
    disp_cols = [
        '상태', p['vn'], p['it'], p['op'], p['vi'], p['av'], 
        '기존리오더', '입고차감', '추가발주', p['t3'], 
        '일판매량', '권장발주수량', '비고(처리내역)'
    ]
    
    # 에디터 시작
    with st.form("final_form"):
        edited_df = st.data_editor(
            df_temp[disp_cols], 
            use_container_width=True, 
            hide_index=True,
            key="main_editor", 
            column_config={
                '상태': st.column_config.TextColumn("상태", disabled=True),
                p['vn']: st.column_config.TextColumn("공급처", disabled=True),
                p['it']: st.column_config.TextColumn("상품명", disabled=True),
                p['op']: st.column_config.TextColumn("옵션", disabled=True),
                p['vi']: st.column_config.TextColumn("공급처명", disabled=True),
                p['av']: st.column_config.NumberColumn("가용재고", disabled=True),
                '기존리오더': st.column_config.NumberColumn("기존리오더", disabled=True),
                '입고차감': st.column_config.NumberColumn("📥 입고(-)", min_value=0), 
                '추가발주': st.column_config.NumberColumn("➕ 발주(+)", min_value=0),
                '권장발주수량': st.column_config.NumberColumn("권장수량", disabled=True),
                '비고(처리내역)': st.column_config.TextColumn("📝 비고(처리내역)")
            }
        )
        btn_save = st.form_submit_button("🚀 최종 데이터 저장 및 시트 전송", use_container_width=True, type="primary")

    if btn_save:
        # 1. 수치(입고 또는 발주)가 입력된 데이터만 추출
        change_list = edited_df[(edited_df['입고차감'] > 0) | (edited_df['추가발주'] > 0)].copy()
        
        if not change_list.empty:
            # ✅ [수정] 저장 전 상품명과 옵션으로 정렬하여 묶어줌 (a-1, a-2 순서)
            change_list = change_list.sort_values(by=[p['it'], p['op']])

            with st.spinner("🚀 상품별로 분류하여 구글 시트 전송 중..."):
                try:
                    sh = get_sheet()
                    ws_qty = sh.worksheet("발주기록")
                    ws_hist = sh.worksheet("히스토리")
                    
                    now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
                    time_short = datetime.now(KST).strftime('%m/%d')
                    
                    rows_qty, rows_hist = [], []

                    for _, r in change_list.iterrows():
                        q_val = int(r['추가발주'])
                        i_val = int(r['입고차감'])
                        user_memo = str(r['비고(처리내역)']).strip() if r['비고(처리내역)'] and str(r['비고(처리내역)']) != "None" else ""
                        
                        # 자동 메모 생성
                        parts = []
                        if q_val > 0: parts.append(f"{q_val}발주")
                        if i_val > 0: parts.append(f"-{i_val}입고")
                        auto_memo = f"[{time_short} " + " ".join(parts) + "]"
                        final_memo = f"{auto_memo} {user_memo}".strip()
                        
                        # 1. 발주기록 시트용 (기존 내부 순서 유지)
                        rows_qty.append([
                            now_s, r[p['it']], r[p['op']], r[p['vi']], r[p['av']], 
                            r['기존리오더'], q_val, r['권장발주수량'], final_memo, r[p['vn']], i_val
                        ])
                        
                        # ✅ 2. 히스토리 시트용 (시간순/상품별 묶음 저장)
                        # 순서: A(시간), B(공급처), C(상품명), D(옵션), E(공급처명), F(가용재고), G(기존리오더), H(입고), I(발주), J(권장수량), K(비고)
                        rows_hist.append([
                            now_s,              # A: 발주시간
                            r[p['vn']],         # B: 공급처(업체명)
                            r[p['it']],         # C: 상품명
                            r[p['op']],         # D: 옵션
                            r[p['vi']],         # E: 공급처상품명
                            r[p['av']],         # F: 가용재고
                            r['기존리오더'],     # G: 기존리오더
                            i_val,              # H: 입고수량
                            q_val,              # I: 추가발주
                            r['권장발주수량'],   # J: 권장수량
                            final_memo          # K: 비고(처리내역)
                        ])
                    
                    # 시트 전송
                    if rows_qty: ws_qty.append_rows(rows_qty)
                    if rows_hist: ws_hist.append_rows(rows_hist)
                    
                    # 5단계 데이터 새로고침
                    if 'db_history' in st.session_state:
                        del st.session_state.db_history
                    
                    st.success(f"✅ 저장 완료! 상품별로 정렬되었습니다. ({len(rows_hist)}건)")
                    time.sleep(1)
                    st.rerun() 
                    
                except Exception as e:
                    st.error(f"저장 실패: {e}")
        else:
            st.warning("⚠️ 입력된 수정 사항(입고 또는 발주)이 없습니다.")
            

# ------------------------------------------------------------------
# 5️⃣단계 & 6️⃣단계: 히스토리 내역 조회 (조회 버튼 통합 버전)
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.header("📜 5단계 & 6단계: 히스토리 내역 조회")

    # 1. 조회 필터 레이아웃
    col_d1, col_d2, col_btn = st.columns([1, 1, 1])
    
    with col_d1:
        # 오늘 날짜를 기본값으로 설정
        start_d = st.date_input("📅 시작일", datetime.now(KST).date(), key="h_start_date")
    with col_d2:
        end_d = st.date_input("📅 종료일", datetime.now(KST).date(), key="h_end_date")
    with col_btn:
        st.write("") # 버튼 위치 조절용 공백
        # 🔍 이 버튼을 눌러야만 실제 조회가 시작됩니다.
        btn_search = st.button("🔍 데이터 조회하기", use_container_width=True, type="primary")

    # 2. 조회 버튼 클릭 시 데이터 로드 로직
    if btn_search:
        with st.spinner("⏳ 해당 기간의 기록을 구글 시트에서 불러오는 중..."):
            try:
                sh = get_sheet()
                ws_hist = sh.worksheet("히스토리")
                raw_data = ws_hist.get_all_values()
                
                if len(raw_data) > 1:
                    # 헤더 정리 및 데이터프레임 생성
                    cols_h = [c.strip() for c in raw_data[0]]
                    df_h = pd.DataFrame(raw_data[1:], columns=cols_h)
                    
                    # '발주시간' 컬럼을 날짜 형식으로 변환 (A열 기준)
                    df_h['발주시간_dt'] = pd.to_datetime(df_h['발주시간'], errors='coerce')
                    
                    # 날짜 필터링 (시작일 <= 데이터 <= 종료일)
                    mask = (df_h['발주시간_dt'].dt.date >= start_d) & (df_h['발주시간_dt'].dt.date <= end_d)
                    filtered_df = df_h.loc[mask].sort_values(by='발주시간_dt', ascending=False)
                    
                    # 세션에 결과 저장
                    st.session_state.db_history = filtered_df
                else:
                    st.session_state.db_history = pd.DataFrame()
                    st.warning("⚠️ 히스토리에 저장된 데이터가 아예 없습니다.")
            except Exception as e:
                st.error(f"❌ 데이터 로드 중 오류 발생: {e}")

    # 3. 결과 출력 및 6단계 상세 검색
    if 'db_history' in st.session_state:
        h_df_display = st.session_state.db_history.copy()
        
        if not h_df_display.empty:
            st.divider()
            st.subheader("🔎 검색 결과 내 상세 필터링 (6단계)")
            
            # 불러온 데이터 내에서 실시간 텍스트 검색 (상품명, 업체명 등)
            h_search = st.text_input("📝 검색어 입력 (공급처, 상품명, 옵션 등)", key="h_sub_search")
            
            if h_search:
                # 모든 컬럼을 문자열로 바꿔서 검색어 포함 여부 확인
                h_df_display = h_df_display[
                    h_df_display.astype(str).apply(lambda x: x.str.contains(h_search, case=False)).any(axis=1)
                ]
            
            st.write(f"✅ 총 **{len(h_df_display)}**건의 기록이 검색되었습니다.")
            
            # 최종 데이터 표 출력 (불필요한 내부 계산용 컬럼은 제외)
            st.dataframe(
                h_df_display.drop(columns=['발주시간_dt'], errors='ignore'), 
                use_container_width=True, 
                hide_index=True
            )
        else:
            if btn_search: # 버튼을 눌렀는데 데이터가 없는 경우만 표시
                st.info("💡 선택하신 기간에는 저장된 히스토리 내역이 없습니다.")
                

# ------------------------------------------------------------------
# 6️⃣단계: 실시간 리오더 현황판 (5단계 조회 데이터 연동)
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.header("📊 6단계: 실시간 리오더 현황판")

    # [중요] 5단계에서 btn_search를 눌러 로드된 'db_history' 데이터를 기반으로 분석합니다.
    if 'db_history' in st.session_state and not st.session_state.db_history.empty:
        m_df = st.session_state.db_history.copy()
        
        # 1. 컬럼 자동 매칭 (다양한 시트 양식 대응)
        qty_col = next((c for c in ['추가발주', '발주', '발주수량'] if c in m_df.columns), '추가발주')
        in_col = next((c for c in ['입고수량', '입고', '입고차감'] if c in m_df.columns), '입고수량')
        date_col = next((c for c in ['발주시간', '날짜', '등록일'] if c in m_df.columns), m_df.columns[0])
        memo_col = next((c for c in ['비고(처리내역)', '메모', '비고'] if c in m_df.columns), '비고(처리내역)')
        v_col = next((c for c in ['공급처', '업체명'] if c in m_df.columns), '공급처')
        vi_col = next((c for c in ['공급처상품명', '공급처명', '매입상품명'] if c in m_df.columns), '공급처상품명')

        # 데이터 전처리 (숫자 변환 및 날짜 처리)
        m_df[qty_col] = pd.to_numeric(m_df[qty_col], errors='coerce').fillna(0)
        m_df[in_col] = pd.to_numeric(m_df[in_col], errors='coerce').fillna(0)
        
        if '발주시간_dt' not in m_df.columns:
            m_df['발주시간_dt'] = pd.to_datetime(m_df[date_col], errors='coerce')
        m_df['날짜_only'] = m_df['발주시간_dt'].dt.date

        # 2. 리오더 집계 로직
        # 상품명, 옵션, 업체별로 그룹화하여 발주합계와 입고합계를 계산합니다.
        group_cols = [c for c in [v_col, '상품명', '옵션', vi_col] if c in m_df.columns]
        
        # 일자별로 먼저 요약 (사장님 요청하신 [04/14 5발주] 형식 유지)
        daily_summary = m_df.groupby(group_cols + ['날짜_only']).agg({
            qty_col: 'sum',
            in_col: 'sum',
            memo_col: lambda x: " ".join(dict.fromkeys([str(i).strip() for i in x if str(i).strip() and str(i).lower() != 'nan']))
        }).reset_index()

        def format_daily_text(row):
            d_str = row['날짜_only'].strftime('%m/%d')
            q_val, i_val = int(row[qty_col]), int(row[in_col])
            u_memo = str(row[memo_col]).strip()
            # 자동 생성된 날짜 대괄호 제거 후 순수 메모만 추출
            clean_memo = u_memo.split(']')[-1].strip() if ']' in u_memo else u_memo
            
            parts = []
            if q_val > 0: parts.append(f"{q_val}발주")
            if i_val > 0: parts.append(f"-{i_val}입고")
            if not parts: return ""
            res = f"[{d_str} " + " ".join(parts) + "]"
            if clean_memo: res += f" {clean_memo}"
            return res

        daily_summary['일자별메모'] = daily_summary.apply(format_daily_text, axis=1)

        # 3. 최종 합산 및 리오더 잔량 계산
        summary = daily_summary.groupby(group_cols).agg({
            '일자별메모': lambda x: " / ".join([i for i in x if i]),
            '날짜_only': 'max', 
            qty_col: 'sum', 
            in_col: 'sum'
        }).reset_index()
        
        summary['리오더 잔량'] = summary[qty_col] - summary[in_col]
        
        # 잔량이 남아있는 상품만 표시 (중요)
        summary = summary[summary['리오더 잔량'] > 0].sort_values('리오더 잔량', ascending=False)

        # 4. 화면 출력
        if not summary.empty:
            summary.rename(columns={
                '날짜_only': '최근기록일', 
                qty_col: '총발주수량', 
                in_col: '총입고수량', 
                '일자별메모': '비고(처리내역)'
            }, inplace=True)

            st.write(f"🚩 현재 미입고된 리오더 상품이 **{len(summary)}**건 있습니다.")
            
            display_cols = ['최근기록일', v_col, '상품명', '옵션', vi_col, '총발주수량', '총입고수량', '리오더 잔량', '비고(처리내역)']
            st.dataframe(
                summary[[c for c in display_cols if c in summary.columns]], 
                use_container_width=True, 
                hide_index=True
            )
        else:
            st.success("✅ 현재 모든 리오더가 입고 완료되었거나 미입고 내역이 없습니다.")
    else:
        st.info("💡 5단계에서 [조회하기] 버튼을 눌러 데이터를 먼저 불러와주세요.")
