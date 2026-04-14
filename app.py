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

# 1️⃣단계: 파일 업로드
st.header("1️⃣ 파일 업로드 및 데이터 로드")
up_file = st.file_uploader("엑셀 파일을 업로드하세요.", type=['xlsx', 'xls'])

if st.button("🔄 현재 화면 데이터 초기화", use_container_width=True):
    st.session_state.clear()
    st.rerun()

if up_file:
    if 'df_raw' not in st.session_state:
        st.session_state.df_raw = pd.read_excel(up_file)
        st.session_state.analyzed = False
        st.success("✅ 파일 로드 완료!")

# 2️⃣~3️⃣단계: 매핑 및 분석 실행
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

# ------------------------------------------------------------------
    # 3️⃣단계: 분석 설정 및 실행 (컬럼명 '입고수량' 확정 버전)
    # ------------------------------------------------------------------
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

        with st.spinner("📊 발주/입고 시트 대조 및 잔량 분석 중..."):
            try:
                df = st.session_state.df_raw.copy()
                today = datetime.now(KST).date()
                sh = get_sheet()
                
                # [함수] 데이터 로드 및 정리 (공백/중복 제거)
                def get_clean_df(name):
                    ws = sh.worksheet(name)
                    data = ws.get_all_values()
                    if len(data) > 1:
                        res = pd.DataFrame(data[1:], columns=[c.strip() for c in data[0]])
                        return res.loc[:, ~res.columns.duplicated()]
                    return pd.DataFrame()

                # 시트 로드
                df_qty = get_clean_df("발주기록")
                df_in = get_clean_df("입고기록")
                st.session_state.master_log = df_qty # 5, 6단계용 저장

                r_map = {}
                # 1. 발주기록에서 총 발주량 합산
                if not df_qty.empty:
                    # 사진에서 확인한 사장님 시트 컬럼명 적용
                    it_c, op_c, q_c = '상품명', '옵션', '추가발주'
                    
                    if it_c in df_qty.columns and op_c in df_qty.columns and q_c in df_qty.columns:
                        df_qty[q_c] = pd.to_numeric(df_qty[q_c], errors='coerce').fillna(0)
                        qty_sum = df_qty.groupby([it_c, op_c])[q_c].sum()

                        # 2. 입고기록에서 총 입고량 합산 ('입고수량' 기준)
                        if not df_in.empty and '입고수량' in df_in.columns:
                            df_in['입고수량'] = pd.to_numeric(df_in['입고수량'], errors='coerce').fillna(0)
                            in_sum = df_in.groupby([it_c, op_c])['입고수량'].sum()
                            # 인덱스 이름 통일 (에러 방지 핵심)
                            in_sum.index.names = [it_c, op_c]
                        else:
                            # 입고 데이터가 없으면 모두 0으로 처리
                            in_sum = pd.Series(0, index=qty_sum.index)

                        # 3. 잔량 계산 (발주량 - 입고량)
                        final_res = qty_sum.sub(in_sum, fill_value=0).clip(lower=0)
                        r_map = final_res.to_dict()

                # 4. 분석 계산 로직 적용
                df[avail] = pd.to_numeric(df[avail], errors='coerce').fillna(0).astype(int)
                df[t1w] = pd.to_numeric(df[t1w], errors='coerce').fillna(0).astype(int)
                
                def get_reorder_val(row):
                    # 현재 분석중인 파일의 상품명/옵션과 시트의 잔량을 매칭
                    k = (str(row[item]).strip(), str(row[option]).strip())
                    return int(r_map.get(k, 0))
                
                # '기존리오더' 칸에 실시간 잔량 주입
                df['기존리오더'] = df.apply(get_reorder_val, axis=1)

                # 판매량 및 권장발주수량 최종 계산
                def get_daily_avg(row):
                    try:
                        r_dt = pd.to_datetime(row[reg_date]).date()
                        days = max(1, min((today - r_dt).days, 7))
                        return int(round(pd.to_numeric(row[t1w]) / days, 0))
                    except: return int(round(pd.to_numeric(row[t1w]) / 7, 0))

                df['일판매량'] = df.apply(get_daily_avg, axis=1)
                df['권장발주수량'] = ((df['일판매량'] * (lt + ss)) - (df[avail] + df['기존리오더'])).clip(lower=0).astype(int)
                
                # 상태 및 4단계용 컬럼 세팅
                df['상태'] = df.apply(lambda r: "🚫 품절" if "품절" in str(r[sold_out]) else ("🚨 발주필요" if r['권장발주수량'] > 0 else "✅ 정상"), axis=1)
                df['입고차감'] = 0
                df['추가발주'] = 0
                df['비고(메모)'] = ""
                
                st.session_state.df_final = df
                st.session_state.analyzed = True
                st.success("✅ 분석 완료! 시트의 '입고수량'을 대조하여 잔량이 반영되었습니다.")
                st.rerun()
                
            except Exception as e:
                st.error(f"⚠️ 분석 오류: {e}")
                st.info("💡 입고기록 시트의 제목을 '입고수량'으로 변경하셨는지 확인해주세요.")

                
# ------------------------------------------------------------------
# 4️⃣단계: 입고 관리 및 최종 저장 (데이터 유실 방지 및 수정 불가 설정)
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.header("📊 4단계: 입고 관리 및 최종 발주 확정")
    
    p = st.session_state.p
    
    # [중요] 검색 시 데이터 날아가는 문제 방지를 위해 st.session_state.df_final을 직접 사용
    if 'edited_df_state' not in st.session_state:
        st.session_state.edited_df_state = st.session_state.df_final.copy()

    f1, f2 = st.columns([1, 2])
    with f1: f_mode = st.selectbox("🚦 상태 필터", ["전체보기", "🚨 발주필요(세트)", "✅ 정상", "🚫 품절"], index=1)
    with f2: s_query = st.text_input("🔍 검색 (상품명/옵션)")

    # 필터링용 임시 DF
    df_temp = st.session_state.df_final.copy()
    if f_mode == "🚨 발주필요(세트)":
        need_items = df_temp[(df_temp['상태'] != "🚫 품절") & (df_temp['권장발주수량'] > 0)][p['it']].unique()
        df_temp = df_temp[df_temp[p['it']].isin(need_items)]
    elif f_mode != "전체보기":
        df_temp = df_temp[df_temp['상태'] == f_mode]
    if s_query:
        df_temp = df_temp[df_temp[p['it']].str.contains(s_query, case=False) | df_temp[p['op']].str.contains(s_query, case=False)]

    disp_cols = ['상태', p['vn'], p['it'], p['op'], p['vi'], p['av'], '기존리오더', '입고차감', '추가발주', p['t3'], '일판매량', '권장발주수량', '비고(메모)']
    
    # 에디터 시작
    with st.form("final_form"):
        # edited_df를 바로 세션에 저장하여 검색을 바꿔도 데이터가 유지되게 함
        edited_df = st.data_editor(
            df_temp[disp_cols], 
            use_container_width=True, 
            hide_index=True,
            key="main_editor", # 키를 지정하여 상태 유지
            column_config={
                '상태': st.column_config.TextColumn("상태", disabled=True),
                p['vn']: st.column_config.TextColumn("공급처", disabled=True),
                p['it']: st.column_config.TextColumn("상품명", disabled=True),
                p['op']: st.column_config.TextColumn("옵션", disabled=True),
                p['vi']: st.column_config.TextColumn("공급처명", disabled=True),
                p['av']: st.column_config.NumberColumn("가용재고", disabled=True),
                '기존리오더': st.column_config.NumberColumn("기존리오더", disabled=True), # 수정 금지 설정
                '입고차감': st.column_config.NumberColumn("📥 입고(-)", min_value=0),
                '추가발주': st.column_config.NumberColumn("➕ 발주(+)", min_value=0),
                '권장발주수량': st.column_config.NumberColumn("권장수량", disabled=True),
                '비고(메모)': st.column_config.TextColumn("📝 메모")
            }
        )
        btn_save = st.form_submit_button("🚀 최종 데이터 저장 및 시트 전송", use_container_width=True, type="primary")

    if btn_save:
        # data_editor의 변경사항은 세션의 'main_editor' 안에 들어있으므로 이를 반영하여 저장
        # (Streamlit 특성상 form 내부의 data_editor는 바로 change_list를 추출하면 됩니다)
        change_list = edited_df[(edited_df['입고차감'] > 0) | (edited_df['추가발주'] > 0)]
        
        if not change_list.empty:
            try:
                sh = get_sheet()
                ws_in = sh.worksheet("입고기록")
                ws_qty = sh.worksheet("발주기록")
                ws_hist = sh.worksheet("히스토리")
                
                now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
                time_short = datetime.now(KST).strftime('%m/%d')
                
                rows_in, rows_qty, rows_hist = [], [], []

                for _, r in change_list.iterrows():
                    q_val = int(r['추가발주'])
                    i_val = int(r['입고차감'])
                    user_memo = str(r['비고(메모)']).strip() if r['비고(메모)'] and str(r['비고(메모)']) != "None" else ""
                    
                    parts = []
                    if q_val > 0: parts.append(f"{q_val}발주")
                    if i_val > 0: parts.append(f"-{i_val}입고")
                    auto_memo = f"[{time_short} " + " ".join(parts) + "]"
                    final_memo = f"{auto_memo} {user_memo}".strip()
                    
                    if i_val > 0:
                        rows_in.append([now_s, r[p['vn']], r[p['it']], r[p['op']], r[p['vi']], i_val, final_memo])
                    if q_val > 0:
                        rows_qty.append([now_s, r[p['vn']], r[p['it']], r[p['op']], r[p['vi']], q_val, final_memo])
                    
                    rows_hist.append([
                        now_s, r[p['vn']], r[p['it']], r[p['op']], r[p['vi']], 
                        r[p['av']], r['기존리오더'], i_val, q_val, r['권장발주수량'], final_memo
                    ])
                
                if rows_in: ws_in.append_rows(rows_in)
                if rows_qty: ws_qty.append_rows(rows_qty)
                if rows_hist: ws_hist.append_rows(rows_hist)
                
                st.success(f"✅ 전송 완료! (입고 {len(rows_in)} / 발주 {len(rows_qty)} / 히스토리 {len(rows_hist)})")
                time.sleep(1)
                st.rerun() # 다시 시작하면서 3단계를 거치면 실시간 잔량이 다시 계산됩니다.
            except Exception as e:
                st.error(f"저장 실패: {e}")
                

# ------------------------------------------------------------------
# 5️⃣단계: 전체 히스토리 기록 (필터 복구 버전)
# ------------------------------------------------------------------
st.divider()
st.header("📜 5단계: 전체 히스토리 기록")

if st.button("🔄 히스토리 시트 불러오기", use_container_width=True, key="h_refresh_final_v13"):
    try:
        sh = get_sheet()
        ws_hist = sh.worksheet("히스토리")
        raw_data = ws_hist.get_all_values()
        
        if len(raw_data) > 1:
            h_df = pd.DataFrame(raw_data[1:], columns=[c.strip() for c in raw_data[0]])
            h_df = h_df.loc[:, ~h_df.columns.duplicated()]
            st.session_state.db_history = h_df
            st.success("✅ 히스토리 데이터를 가져왔습니다.")
            st.rerun()
        else:
            st.warning("⚠️ 히스토리 시트에 데이터가 없습니다.")
    except Exception as e:
        st.error(f"히스토리 로드 실패: {e}")

if 'db_history' in st.session_state and not st.session_state.db_history.empty:
    m_df_5 = st.session_state.db_history.copy()
    date_col_5 = next((c for c in m_df_5.columns if '날짜' in c), m_df_5.columns[0])
    m_df_5['날짜_dt'] = pd.to_datetime(m_df_5[date_col_5], errors='coerce')
    m_df_5['날짜_only'] = m_df_5['날짜_dt'].dt.date
    valid_df_5 = m_df_5.dropna(subset=['날짜_dt']).copy()
    
    if not valid_df_5.empty:
        c1, c2, c3 = st.columns(3)
        with c1: 
            sel_dates_5 = st.date_input("📅 조회 날짜 범위", [valid_df_5['날짜_only'].min(), valid_df_5['날짜_only'].max()], key="h_date_v13")
        with c2: 
            h_name_5 = st.text_input("🔍 상품명 검색", key="h_name_v13")
        with c3: 
            t_options_5 = ["전체 회차"] + sorted(valid_df_5['날짜_dt'].dt.strftime('%Y-%m-%d %H:%M:%S').unique(), reverse=True)
            h_time_5 = st.selectbox("⏰ 저장 회차 선택", t_options_5, key="h_time_v13")

        df_display_5 = valid_df_5.copy()
        if isinstance(sel_dates_5, (list, tuple)) and len(sel_dates_5) == 2:
            df_display_5 = df_display_5[(df_display_5['날짜_only'] >= sel_dates_5[0]) & (df_display_5['날짜_only'] <= sel_dates_5[1])]
        if h_name_5:
            df_display_5 = df_display_5[df_display_5['상품명'].astype(str).str.contains(h_name_5, case=False)]
        if h_time_5 != "전체 회차":
            df_display_5 = df_display_5[df_display_5['날짜_dt'].dt.strftime('%Y-%m-%d %H:%M:%S') == h_time_5]

        st.dataframe(df_display_5.sort_values('날짜_dt', ascending=False), use_container_width=True, hide_index=True)

# ------------------------------------------------------------------
# 6️⃣단계: 실시간 리오더 현황판 (KeyError 방지 유연한 매칭 버전)
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.header("📊 6단계: 실시간 리오더 현황판")

    if st.button("🔄 실시간 장부 데이터 동기화", use_container_width=True, type="secondary", key="sync_v13"):
        with st.spinner("📡 '발주기록' 시트 분석 중..."):
            try:
                sh = get_sheet()
                ws_log = sh.worksheet("발주기록")
                data = ws_log.get_all_values()
                if len(data) > 1:
                    df_refresh = pd.DataFrame(data[1:], columns=[c.strip() for c in data[0]])
                    st.session_state.master_log = df_refresh
                    st.success("✅ 최신 장부 데이터 동기화 완료!")
                    st.rerun()
            except Exception as e:
                st.error(f"데이터 로드 실패: {e}")

    if 'master_log' in st.session_state and not st.session_state.master_log.empty:
        m_df = st.session_state.master_log.copy()
        
        # [중요] KeyError 방지를 위해 실제 시트의 컬럼명을 유연하게 찾습니다.
        qty_col = next((c for c in ['추가발주', '발주수량', '추가발주수량'] if c in m_df.columns), None)
        in_col = next((c for c in ['입고수량', '입고'] if c in m_df.columns), None)
        date_col = next((c for c in ['날짜', '등록일'] if c in m_df.columns), '날짜')
        memo_col = next((c for c in ['메모', '비고'] if c in m_df.columns), '메모')
        v_col = next((c for c in ['업체명', '공급처'] if c in m_df.columns), '업체명')
        vi_col = next((c for c in ['공급처상품명', '매입상품명'] if c in m_df.columns), '공급처상품명')

        # 필수 컬럼 존재 확인
        if not qty_col or not in_col:
            st.error(f"⚠️ 시트에서 '추가발주' 또는 '입고수량' 컬럼을 찾을 수 없습니다. (현재 컬럼: {list(m_df.columns)})")
        else:
            m_df['날짜_dt'] = pd.to_datetime(m_df[date_col], errors='coerce')
            m_df['날짜_only'] = m_df['날짜_dt'].dt.date
            m_df = m_df.dropna(subset=['날짜_dt'])

            # --- [6단계 상단 필터] ---
            f1, f2, f3 = st.columns([1.5, 1.5, 1])
            with f1:
                r_date = st.date_input("📅 현황 확인 기간", [m_df['날짜_only'].min(), m_df['날짜_only'].max()], key="r_date_v13")
            with f2:
                r_name = st.text_input("🔍 상품명 검색", key="r_name_v13")
            with f3:
                v_list = ["전체 업체"] + sorted([str(v) for v in m_df[v_col].unique() if str(v).strip()]) if v_col in m_df.columns else ["전체 업체"]
                r_vendor = st.selectbox("🏭 업체 필터", v_list, key="r_vendor_v13")

            df_6 = m_df.copy()
            if isinstance(r_date, (list, tuple)) and len(r_date) == 2:
                df_6 = df_6[(df_6['날짜_only'] >= r_date[0]) & (df_6['날짜_only'] <= r_date[1])]
            if r_name:
                df_6 = df_6[df_6['상품명'].astype(str).str.contains(r_name, case=False)]
            if v_col in df_6.columns and r_vendor != "전체 업체":
                df_6 = df_6[df_6[v_col] == r_vendor]

            if not df_6.empty:
                # 숫자 변환
                df_6[qty_col] = pd.to_numeric(df_6[qty_col], errors='coerce').fillna(0)
                df_6[in_col] = pd.to_numeric(df_6[in_col], errors='coerce').fillna(0)
                
                group_cols = [c for c in [v_col, '상품명', '옵션', vi_col, '날짜_only'] if c in df_6.columns]
                
                # 일자별 집계
                daily_summary = df_6.groupby(group_cols).agg({
                    qty_col: 'sum',
                    in_col: 'sum',
                    memo_col: lambda x: " ".join(dict.fromkeys([str(i).strip() for i in x if str(i).strip() and str(i).lower() != 'nan']))
                }).reset_index()

                def format_daily_text(row):
                    d_str = row['날짜_only'].strftime('%m/%d')
                    q_val, i_val = int(row[qty_col]), int(row[in_col])
                    u_memo = str(row[memo_col]).replace('nan', '').strip()
                    parts = []
                    if q_val > 0: parts.append(f"{q_val}발주")
                    if i_val > 0: parts.append(f"-{i_val}입고")
                    if not parts: return ""
                    res = f"[{d_str} " + " ".join(parts) + "]"
                    if u_memo: res += f" {u_memo}"
                    return res

                daily_summary['일자별메모'] = daily_summary.apply(format_daily_text, axis=1)

                final_group = [k for k in group_cols if k != '날짜_only']
                summary = daily_summary.groupby(final_group).agg({
                    '일자별메모': lambda x: " / ".join([i for i in x if i]),
                    '날짜_only': 'max', qty_col: 'sum', in_col: 'sum'
                }).reset_index()
                
                summary['리오더 잔량'] = summary[qty_col] - summary[in_col]
                summary = summary[summary['리오더 잔량'] > 0].sort_values('리오더 잔량', ascending=False)

                summary.rename(columns={'날짜_only': '최근기록일', qty_col: '발주수량', in_col: '입고수량', '일자별메모': '비고(처리내역)'}, inplace=True)
                summary['총발주량'] = summary['발주수량'] 

                display_order = ['최근기록일', v_col, '상품명', '옵션', vi_col, '총발주량', '입고수량', '발주수량', '리오더 잔량', '비고(처리내역)']
                st.dataframe(summary[[c for c in display_order if c in summary.columns]], use_container_width=True, hide_index=True)
            else:
                st.info("조회된 데이터가 없습니다.")
