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
# 4️⃣단계: 입고 관리 및 최종 저장 (순서교정 + 실시간계산 + 마이너스방지 완결판)
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.header("📊 4단계: 입고 관리 및 최종 발주 확정")
    
    p = st.session_state.p
    
    # 데이터 체크
    if 'df_final' not in st.session_state:
        st.error("데이터가 없습니다. 1~3단계를 먼저 진행해주세요.")
        st.stop()

    f1, f2 = st.columns([1, 2])
    with f1: f_mode = st.selectbox("🚦 상태 필터", ["전체보기", "🚨 발주필요(세트)", "✅ 정상", "🚫 품절"], index=1)
    with f2: s_query = st.text_input("🔍 검색 (상품명/옵션)")

    # 화면 표시용 필터링
    df_temp = st.session_state.df_final.copy()
    if f_mode == "🚨 발주필요(세트)":
        need_items = df_temp[(df_temp['상태'] != "🚫 품절") & (df_temp['권장발주수량'] > 0)][p['it']].unique()
        df_temp = df_temp[df_temp[p['it']].isin(need_items)]
    elif f_mode != "전체보기":
        df_temp = df_temp[df_temp['상태'] == f_mode]
    if s_query:
        df_temp = df_temp[df_temp[p['it']].str.contains(s_query, case=False) | df_temp[p['op']].str.contains(s_query, case=False)]

    disp_cols = [
        '상태', p['vn'], p['it'], p['op'], p['vi'], p['av'], 
        '기존리오더', '입고차감', '추가발주', p['t3'], 
        '일판매량', '권장발주수량', '비고(처리내역)'
    ]
    
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
        # 변경 사항이 있는 데이터만 추출
        changed_rows = edited_df[(edited_df['입고차감'] > 0) | (edited_df['추가발주'] > 0)].copy()
        
        if not changed_rows.empty:
            changed_rows = changed_rows.sort_values(by=[p['it'], p['op']])

            with st.spinner("🚀 수치 계산 및 시트 전송 중..."):
                try:
                    sh = get_sheet()
                    ws_qty = sh.worksheet("발주기록")
                    ws_hist = sh.worksheet("히스토리")
                    
                    now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
                    time_short = datetime.now(KST).strftime('%m/%d')
                    
                    rows_qty, rows_hist = [], []

                    for idx, r in changed_rows.iterrows():
                        q_val = int(r['추가발주'])
                        i_val = int(r['입고차감'])
                        user_memo = str(r['비고(처리내역)']).strip() if r['비고(처리내역)'] and str(r['비고(처리내역)']) != "None" else ""
                        
                        # 자동 메모 생성
                        parts = []
                        if q_val > 0: parts.append(f"{q_val}발주")
                        if i_val > 0: parts.append(f"-{i_val}입고")
                        auto_memo = f"[{time_short} " + " ".join(parts) + "]"
                        final_memo = f"{auto_memo} {user_memo}".strip()
                        
                        # 1. [발주기록 시트] 저장 순서 교정 (A~J열)
                        rows_qty.append([
                            now_s, r[p['vn']], r[p['it']], r[p['op']], r[p['vi']], 
                            r[p['av']], r['기존리오더'], q_val, i_val, final_memo
                        ])
                        
                        # 2. [히스토리 시트] 저장
                        rows_hist.append([
                            now_s, r[p['vn']], r[p['it']], r[p['op']], r[p['vi']], 
                            r[p['av']], r['기존리오더'], i_val, q_val, r['권장발주수량'], final_memo
                        ])

                        # 3. [실시간 계산] 화면 데이터(session_state) 업데이트
                        mask = (st.session_state.df_final[p['it']] == r[p['it']]) & (st.session_state.df_final[p['op']] == r[p['op']])
                        
                        # 리오더 계산: 기존 + 발주 - 입고
                        new_reorder = r['기존리오더'] + q_val - i_val
                        
                        # ✨ 마이너스 방지: 0보다 작으면 0으로 고정
                        st.session_state.df_final.loc[mask, '기존리오더'] = max(0, new_reorder)
                        
                        # ✨ 입력값 리셋 (0으로 초기화)
                        st.session_state.df_final.loc[mask, '추가발주'] = 0
                        st.session_state.df_final.loc[mask, '입고차감'] = 0
                        st.session_state.df_final.loc[mask, '비고(처리내역)'] = ""

                    # 시트 실제 전송
                    if rows_qty: ws_qty.append_rows(rows_qty, value_input_option='USER_ENTERED')
                    if rows_hist: ws_hist.append_rows(rows_hist, value_input_option='USER_ENTERED')
                    
                    # 캐시 삭제 (6단계 연동용)
                    if 'master_log' in st.session_state: del st.session_state.master_log
                    if 'db_history' in st.session_state: del st.session_state.db_history
                    
                    st.success(f"✅ 저장 완료 및 수치 업데이트 성공! ({len(rows_hist)}건)")
                    time.sleep(1)
                    st.rerun() # 화면 새로고침하여 0으로 리셋된 결과 보여줌
                    
                except Exception as e:
                    st.error(f"저장 실패: {e}")
        else:
            st.warning("⚠️ 입력된 수정 사항(입고 또는 발주)이 없습니다.")
            

# ------------------------------------------------------------------
# 5️⃣단계: 전체 히스토리 기록 (조회 버튼 위치 및 정렬 수정)
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.header("📜 5단계: 전체 히스토리 기록")

    if 'db_history' not in st.session_state:
        try:
            sh = get_sheet()
            ws_hist = sh.worksheet("히스토리")
            raw_data = ws_hist.get_all_values()
            if len(raw_data) > 1:
                cols_5 = [c.strip() for c in raw_data[0]]
                h_df = pd.DataFrame(raw_data[1:], columns=cols_5)
                # 중복 컬럼 제거 및 명칭 통일
                h_df = h_df.loc[:, ~h_df.columns.duplicated()]
                h_df.rename(columns={'메모': '비고(처리내역)', '비고': '비고(처리내역)', '비고(메모)': '비고(처리내역)'}, errors='ignore', inplace=True)
                st.session_state.db_history = h_df
            else:
                st.session_state.db_history = pd.DataFrame()
        except:
            st.session_state.db_history = pd.DataFrame()

    m_df_5 = st.session_state.get('db_history', pd.DataFrame()).copy()
    
    if not m_df_5.empty:
        # 날짜 컬럼 파싱 (정렬을 위해 필수)
        d_col = next((c for c in m_df_5.columns if '날짜' in c or '시간' in c), m_df_5.columns[0])
        m_df_5['날짜_dt'] = pd.to_datetime(m_df_5[d_col], errors='coerce')
        m_df_5['날짜_only'] = m_df_5['날짜_dt'].dt.date
        
        # 🛠️ 버튼 위치 조정: 날짜 범위 바로 옆으로 배치
        c1, c2, c3, c4 = st.columns([1.2, 0.4, 1.2, 1.2]) # c2에 버튼 배치
        with c1: 
            sel_dates_5 = st.date_input("📅 조회 날짜 범위", [m_df_5['날짜_only'].min(), m_df_5['날짜_only'].max()], key="h_date_v15")
        with c2: 
            st.write("") # 라벨 높이 맞춤용
            btn_h_run = st.button("🔍 조회", key="btn_h_run", use_container_width=True)
        with c3: 
            h_name_5 = st.text_input("🔍 상품명 검색", key="h_name_v15")
        with c4:
            t_opts = ["전체 회차"] + sorted(m_df_5['날짜_dt'].dropna().dt.strftime('%Y-%m-%d %H:%M:%S').unique(), reverse=True)
            h_time_5 = st.selectbox("⏰ 저장 회차 선택", t_opts, key="h_time_v15")

        # 필터링 및 정렬 로직
        df_dis = m_df_5.copy()
        
        if len(sel_dates_5) == 2:
            df_dis = df_dis[(df_dis['날짜_only'] >= sel_dates_5[0]) & (df_dis['날짜_only'] <= sel_dates_5[1])]
        if h_name_5:
            df_dis = df_dis[df_dis.apply(lambda r: h_name_5.lower() in str(r).lower(), axis=1)]
        if h_time_5 != "전체 회차":
            df_dis = df_dis[df_dis['날짜_dt'].dt.strftime('%Y-%m-%d %H:%M:%S') == h_time_5]

        # ✨ 정렬 강화: 최신 발주시간이 무조건 위로 오도록 정렬
        df_dis = df_dis.sort_values(by='날짜_dt', ascending=False)

        st.dataframe(
            df_dis.drop(columns=['날짜_dt', '날짜_only'], errors='ignore'), 
            use_container_width=True, 
            hide_index=True
        )
    else:
        st.info("💡 히스토리 내역이 없습니다. 4단계에서 데이터를 먼저 저장해주세요.")



# ------------------------------------------------------------------
# 6️⃣단계: 실시간 리오더 현황판 (에러 교정 및 검색 UI 완결판)
# ------------------------------------------------------------------
def render_step6():
    st.divider()
    st.header("📈 6단계: 실시간 리오더 현황판")
    
    # 1. 데이터 로드 (세션 캐시 사용)
    if 'master_log' not in st.session_state:
        try:
            sh = get_sheet()
            ws_qty = sh.worksheet("발주기록")
            raw_data = ws_qty.get_all_records()
            st.session_state.master_log = pd.DataFrame(raw_data)
        except Exception as e:
            st.error(f"데이터 로드 실패: {e}")
            return

    df_log = st.session_state.master_log.copy()

    if not df_log.empty:
        # 🚨 [핵심 수정] 모든 계산용 컬럼을 강제로 숫자형으로 변환 (에러 원천 차단)
        num_cols = ['추가발주', '입고수량', '기존리오더']
        for col in num_cols:
            if col in df_log.columns:
                # 숫자가 아닌 값은 NaN으로 바꾸고, NaN은 0으로 채움
                df_log[col] = pd.to_numeric(df_log[col], errors='coerce').fillna(0)

        # 2. 상단 필터 및 검색 UI (다시 살려냈습니다!)
        c1, c2, c3 = st.columns([1, 1, 1])
        with c1:
            v_list = ["전체"] + sorted(df_log['공급처'].unique().tolist())
            sel_v = st.selectbox("🏭 공급처 필터", v_list, key="step6_v_filter")
        with c2:
            sel_s = st.text_input("🔍 상품명 검색", key="step6_s_search")
        with c3:
            st.write("") # 간격 조정
            if st.button("🔄 현황판 새로고침", use_container_width=True):
                if 'master_log' in st.session_state:
                    del st.session_state.master_log
                st.rerun()

        # 데이터 필터링 적용
        df_dash = df_log.copy()
        if sel_v != "전체":
            df_dash = df_dash[df_dash['공급처'] == sel_v]
        if sel_s:
            df_dash = df_dash[df_dash['상품명'].str.contains(sel_s, case=False)]

        # 3. 실시간 집계 로직
        summary = df_dash.groupby(['공급처', '상품명', '옵션']).agg({
            '추가발주': 'sum',
            '입고수량': 'sum',
            '기존리오더': 'first'
        }).reset_index()
        
        # 계산부 (안전하게 float 형변환 후 계산)
        summary['총발주'] = summary['기존리오더'].astype(float) + summary['추가발주'].astype(float)
        # 🚨 마이너스 방지: 계산 결과가 0보다 작으면 0으로 고정
        summary['리오더잔량'] = (summary['총발주'] - summary['입고수량'].astype(float)).apply(lambda x: max(0, x))

        # 4. 상단 요약 메트릭
        m1, m2, m3 = st.columns(3)
        m1.metric("📦 누적 총 발주", f"{int(summary['총발주'].sum())}개")
        m2.metric("📥 누적 총 입고", f"{int(summary['입고수량'].sum())}개")
        m3.metric("⏳ 미입고 잔량", f"{int(summary['리오더잔량'].sum())}개", delta_color="inverse")
        
        # 5. 상세 현황 데이터 테이블
        st.subheader("📝 상세 리오더 현황 리스트")
        st.dataframe(
            summary[['공급처', '상품명', '옵션', '총발주', '입고수량', '리오더잔량']], 
            use_container_width=True, 
            hide_index=True,
            column_config={
                "총발주": "총 발주량",
                "입고수량": "입고 완료",
                "리오더잔량": st.column_config.NumberColumn("남은 잔량", format="%d")
            }
        )
    else:
        st.info("시트에 데이터가 없습니다. 발주를 먼저 진행해 주세요.")
