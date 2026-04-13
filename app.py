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
    """3단계 분석 시 리오더 잔량을 가져오는 함수"""
    r_map = {}
    sh = get_sheet()
    if sh:
        try:
            ws = sh.worksheet("발주기록")
            logs = ws.get_all_values()
            if len(logs) > 1:
                df_l = pd.DataFrame(logs[1:], columns=[c.strip() for c in logs[0]])
                # 수량 컬럼은 7번째(인덱스 6)라고 가정
                df_l['k'] = df_l.apply(lambda r: super_clean(r.iloc[1]) + super_clean(r.iloc[2]), axis=1)
                r_map = df_l.groupby('k').apply(lambda x: x.iloc[:, 6].apply(to_i).sum()).to_dict()
        except: pass
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

    st.divider()
    st.subheader("⚙️ 3️⃣단계: 분석 설정")
    clt, css = st.columns(2)
    lt, ss = clt.number_input("리드타임 (일)", value=10), css.number_input("안전재고 (일 수)", value=7)

    if st.button("🚀 분석 실행 / 업데이트", type="primary", use_container_width=True):
        st.session_state.p = {
            'so': sold_out, 'it': item, 'op': option, 'vn': vendor, 'vi': v_item_col,
            'av': avail, 't3': t3d, 't7': t1w, 'lt': lt, 'ss': ss, 'rd': reg_date
        }

        with st.spinner("📊 구글 시트 동기화 및 데이터 분석 중..."):
            try:
                df = st.session_state.df_raw.copy()
                r_map = load_reorder_data() # 시트 호출
                today = datetime.now(KST).date()

                # 숫자 변환
                df[avail] = pd.to_numeric(df[avail], errors='coerce').fillna(0).astype(int)
                df[t1w] = pd.to_numeric(df[t1w], errors='coerce').fillna(0).astype(int)

                # 키 생성 및 리오더 매핑
                df['clean_k'] = df.apply(lambda r: super_clean(r[item]) + super_clean(r[option]), axis=1)
                df['기존리오더'] = df['clean_k'].map(r_map).fillna(0).astype(int).clip(lower=0)
                
                # 평균 판매량 계산
                def get_daily_avg(row):
                    try:
                        r_dt = pd.to_datetime(row[reg_date]).date()
                        days = max(1, min((today - r_dt).days, 7))
                        return int(round(to_i(row[t1w]) / days, 0))
                    except: return int(round(to_i(row[t1w]) / 7, 0))

                df['일판매량'] = df.apply(get_daily_avg, axis=1)
                df['권장발주수량'] = ((df['일판매량'] * (lt + ss)) - (df[avail] + df['기존리오더'])).clip(lower=0).astype(int)
                
                def status_check(row):
                    if "품절" in str(row[sold_out]): return "🚫 품절"
                    return "🚨 발주필요" if row['권장발주수량'] > 0 else "✅ 정상"
                df['상태'] = df.apply(status_check, axis=1)
                
                df['입고차감'] = 0
                df['추가발주'] = 0
                df['비고(메모)'] = ""
                
                st.session_state.df_final = df
                st.session_state.analyzed = True
                st.rerun()
            except Exception as e:
                st.error(f"분석 오류: {e}")

# 4️⃣단계: 입고 관리 및 최종 발주 확정
if st.session_state.get('analyzed'):
    st.divider()
    st.header("📊 4단계: 입고 관리 및 최종 발주 확정")
    
    p = st.session_state.p
    df_disp = st.session_state.df_final.copy()
    
    # 필터
    f1, f2 = st.columns([1, 2])
    with f1: 
        f_mode = st.selectbox("🚦 상태 필터", ["전체보기", "🚨 발주필요(세트)", "✅ 정상", "🚫 품절"], index=1)
    with f2: 
        s_query = st.text_input("🔍 검색 (상품명/옵션)")

    if f_mode == "🚨 발주필요(세트)":
        df_active = df_disp[df_disp['상태'] != "🚫 품절"]
        need_items = df_active[df_active['권장발주수량'] > 0][p['it']].unique()
        df_disp = df_active[df_active[p['it']].isin(need_items)]
    elif f_mode == "✅ 정상":
        df_disp = df_disp[df_disp['상태'] == "✅ 정상"]
    elif f_mode == "🚫 품절":
        df_disp = df_disp[df_disp['상태'] == "🚫 품절"]

    if s_query:
        df_disp = df_disp[df_disp[p['it']].astype(str).str.contains(s_query, case=False) | 
                          df_disp[p['op']].astype(str).str.contains(s_query, case=False)]

    disp_cols = ['상태', p['vn'], p['it'], p['op'], p['vi'], p['av'], '기존리오더', '입고차감', '추가발주', p['t3'], '일판매량', '권장발주수량', '비고(메모)']
    
    with st.form("final_form"):
        edited_df = st.data_editor(
            df_disp[disp_cols],
            use_container_width=True, hide_index=True,
            column_config={
                '상태': st.column_config.TextColumn("상태", disabled=True),
                p['vn']: st.column_config.TextColumn("공급처", disabled=True),
                p['it']: st.column_config.TextColumn("상품명", disabled=True),
                '기존리오더': st.column_config.NumberColumn("📦 기존잔량", disabled=True),
                '입고차감': st.column_config.NumberColumn("📥 입고(-)", min_value=0),
                '추가발주': st.column_config.NumberColumn("➕ 발주(+)", min_value=0),
                '권장발주수량': st.column_config.NumberColumn("💡 권장", disabled=True)
            }
        )
        btn_save = st.form_submit_button("💾 데이터 최종 저장 및 시트 전송", use_container_width=True, type="primary")

    if btn_save:
        change_list = edited_df[(edited_df['입고차감'] > 0) | (edited_df['추가발주'] > 0)]
        if not change_list.empty:
            try:
                sh = get_sheet()
                ws_log = sh.worksheet("발주기록")
                now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
                rows = [[now_s, r[p['it']], r[p['op']], r[p['vi']], r[p['av']], r['기존리오더'], int(r['추가발주'])-int(r['입고차감']), r['권장발주수량'], r['비고(메모)'], r[p['vn']]] for _, r in change_list.iterrows()]
                ws_log.append_rows(rows)
                st.success("✅ 저장 완료!")
                st.cache_data.clear()
                time.sleep(1)
                st.rerun()
            except Exception as e: st.error(f"저장 실패: {e}")

# 5️⃣~6️⃣단계: 기록 및 리오더 현황 (무한로딩 방지형)
if st.session_state.get('analyzed'):
    st.divider()
    st.header("🕒 5~6단계: 기록 및 리오더 현황")

    def load_master_data():
        try:
            sh = get_sheet()
            if sh:
                ws = sh.worksheet("발주기록")
                st.session_state.master_log = pd.DataFrame(ws.get_all_records())
                st.success("✅ 동기화 완료!")
        except Exception as e: st.error(f"로드 실패: {e}")

    if st.button("🔄 기록 및 리오더 데이터 불러오기", use_container_width=True):
        load_master_data()

    if 'master_log' in st.session_state and not st.session_state.master_log.empty:
        m_df = st.session_state.master_log.copy()
        m_df['일시_dt'] = pd.to_datetime(m_df['일시'], errors='coerce').dt.date

        with st.expander("📜 5단계: 최근 히스토리", expanded=True):
            h_search = st.text_input("🔍 검색어 입력", key="h_search")
            df_h = m_df.copy()
            if h_search:
                df_h = df_h[df_h['상품명'].astype(str).str.contains(h_search, case=False)]
            st.dataframe(df_h.sort_values('일시', ascending=False).drop(columns=['일시_dt']), use_container_width=True, hide_index=True)

        with st.expander("📊 6단계: 실시간 리오더 잔량", expanded=True):
            df_r = m_df.groupby(['상품명', '옵션'])['수량'].sum().reset_index()
            df_r = df_r[df_r['수량'] > 0].sort_values('수량', ascending=False)
            st.dataframe(df_r, use_container_width=True, hide_index=True)
    else:
        st.info("💡 위 버튼을 눌러야 내역이 나타납니다.")
