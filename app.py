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

# ------------------------------------------------------------------
    # 3️⃣단계: 분석 설정 및 실행
    # ------------------------------------------------------------------
    st.divider()
    st.subheader("⚙️ 3️⃣단계: 분석 설정")
    clt, css = st.columns(2)
    lt, ss = clt.number_input("리드타임 (일)", value=10), css.number_input("안전재고 (일 수)", value=7)

    if st.button("🚀 분석 실행 / 전체 장부 업데이트", type="primary", use_container_width=True):
        st.session_state.p = {
            'so': sold_out, 'it': item, 'op': option, 'vn': vendor, 'vi': v_item_col,
            'av': avail, 't3': t3d, 't7': t1w, 'lt': lt, 'ss': ss, 'rd': reg_date
        }

        with st.spinner("📊 구글 시트 동기화 및 데이터 분석 중..."):
            try:
                # 1. 기초 데이터 로드
                df = st.session_state.df_raw.copy()
                today = datetime.now(KST).date()

                # 2. 구글 시트 장부 로드 (5, 6단계용)
                sh = get_sheet()
                ws_log = sh.worksheet("발주기록")
                df_master = pd.DataFrame(ws_log.get_all_records())
                df_master.columns = [c.strip() for c in df_master.columns]
                st.session_state.master_log = df_master 

                # 3. 기존 리오더 잔량 계산 (r_map)
                q_col = '추가발주수량' if '추가발주수량' in df_master.columns else '추가발주'
                df_master[q_col] = pd.to_numeric(df_master[q_col], errors='coerce').fillna(0)
                r_map = df_master.groupby(['상품명', '옵션'])[q_col].sum().to_dict()

                # 4. 분석 계산 로직
                df[avail] = pd.to_numeric(df[avail], errors='coerce').fillna(0).astype(int)
                df[t1w] = pd.to_numeric(df[t1w], errors='coerce').fillna(0).astype(int)
                
                def get_reorder_val(row):
                    return r_map.get((row[item], row[option]), 0)
                
                df['기존리오더'] = df.apply(get_reorder_val, axis=1).fillna(0).astype(int)

                def get_daily_avg(row):
                    try:
                        r_dt = pd.to_datetime(row[reg_date]).date()
                        days = max(1, min((today - r_dt).days, 7))
                        return int(round(pd.to_numeric(row[t1w]) / days, 0))
                    except: return int(round(pd.to_numeric(row[t1w]) / 7, 0))

                df['일판매량'] = df.apply(get_daily_avg, axis=1)
                df['권장발주수량'] = ((df['일판매량'] * (lt + ss)) - (df[avail] + df['기존리오더'])).clip(lower=0).astype(int)
                
                df['상태'] = df.apply(lambda r: "🚫 품절" if "품절" in str(r[sold_out]) else ("🚨 발주필요" if r['권장발주수량'] > 0 else "✅ 정상"), axis=1)
                
                # 4단계용 기본값 세팅
                df['입고차감'] = 0
                df['추가발주'] = 0
                df['비고(메모)'] = ""
                
                st.session_state.df_final = df
                st.session_state.analyzed = True
                st.success("✅ 분석 및 장부 업데이트 완료!")
                st.rerun()
            except Exception as e:
                st.error(f"분석 오류: {e}")

# ------------------------------------------------------------------
# 4️⃣단계: 입고 관리 및 최종 저장
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.header("📊 4단계: 입고 관리 및 최종 발주 확정")
    
    p = st.session_state.p
    df_disp = st.session_state.df_final.copy()
    
    f1, f2 = st.columns([1, 2])
    with f1: f_mode = st.selectbox("🚦 상태 필터", ["전체보기", "🚨 발주필요(세트)", "✅ 정상", "🚫 품절"], index=1)
    with f2: s_query = st.text_input("🔍 검색 (상품명/옵션)")

    # 필터 적용
    if f_mode == "🚨 발주필요(세트)":
        need_items = df_disp[(df_disp['상태'] != "🚫 품절") & (df_disp['권장발주수량'] > 0)][p['it']].unique()
        df_disp = df_disp[df_disp[p['it']].isin(need_items)]
    elif f_mode != "전체보기":
        df_disp = df_disp[df_disp['상태'] == f_mode]
    if s_query:
        df_disp = df_disp[df_disp[p['it']].str.contains(s_query, case=False) | df_disp[p['op']].str.contains(s_query, case=False)]

    disp_cols = ['상태', p['vn'], p['it'], p['op'], p['vi'], p['av'], '기존리오더', '입고차감', '추가발주', p['t3'], '일판매량', '권장발주수량', '비고(메모)']
    
    with st.form("final_form"):
        edited_df = st.data_editor(df_disp[disp_cols], use_container_width=True, hide_index=True,
            column_config={
                '상태': st.column_config.TextColumn("상태", disabled=True),
                p['vn']: st.column_config.TextColumn("공급처", disabled=True),
                p['it']: st.column_config.TextColumn("상품명", disabled=True),
                '입고차감': st.column_config.NumberColumn("📥 입고(-)", min_value=0),
                '추가발주': st.column_config.NumberColumn("➕ 발주(+)", min_value=0),
                '비고(메모)': st.column_config.TextColumn("📝 메모")
            })
        btn_save = st.form_submit_button("💾 데이터 최종 저장 및 시트 전송", use_container_width=True, type="primary")

    if btn_save:
        change_list = edited_df[(edited_df['입고차감'] > 0) | (edited_df['추가발주'] > 0)]
        if not change_list.empty:
            try:
                sh = get_sheet()
                ws_log = sh.worksheet("발주기록")
                now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
                rows = []
                for _, r in change_list.iterrows():
                    # 메모 및 수량 계산
                    memo_parts = []
                    if r['입고차감'] > 0: memo_parts.append(f"-{int(r['입고차감'])} 입고")
                    if r['추가발주'] > 0: memo_parts.append(f"{int(r['추가발주'])} 발주")
                    auto_memo = " / ".join(memo_parts)
                    user_memo = str(r['비고(메모)']).strip() if r['비고(메모)'] and str(r['비고(메모)']) != "None" else ""
                    final_memo = f"[{auto_memo}] {user_memo}" if user_memo else auto_memo
                    
                    # 시트 컬럼 순서 맞춤
                    rows.append([now_s, r[p['vn']], r[p['it']], r[p['op']], r[p['vi']], r[p['av']], r['기존리오더'], int(r['입고차감']), int(r['추가발주']), r['권장발주수량'], final_memo])
                
                ws_log.append_rows(rows)
                st.success("✅ 저장 완료!")
                if 'master_log' in st.session_state: del st.session_state.master_log
                time.sleep(1)
                st.rerun()
            except Exception as e: st.error(f"저장 실패: {e}")

# ------------------------------------------------------------------
# 5️⃣단계: 전체 히스토리 기록 (상단 필터 포함)
# ------------------------------------------------------------------
if st.session_state.get('analyzed') and 'master_log' in st.session_state:
    st.divider()
    st.header("📜 5단계: 전체 히스토리 기록")
    m_df = st.session_state.master_log.copy()
    m_df['날짜'] = pd.to_datetime(m_df['날짜'], errors='coerce')
    m_df['날짜_only'] = m_df['날짜'].dt.date
    
    c1, c2, c3 = st.columns(3)
    with c1: sel_dates = st.date_input("📅 조회 날짜", [m_df['날짜_only'].max() - timedelta(days=7), m_df['날짜_only'].max()], key="h_date")
    with c2: h_name = st.text_input("🔍 상품명 검색", key="h_name")
    with c3: 
        t_options = ["전체 회차"] + [t.strftime('%Y-%m-%d %H:%M:%S') for t in sorted(m_df['날짜'].unique(), reverse=True)]
        h_time = st.selectbox("⏰ 회차(시간) 선택", t_options, key="h_time")

    df_5 = m_df.copy()
    if len(sel_dates) == 2: df_5 = df_5[(df_5['날짜_only'] >= sel_dates[0]) & (df_5['날짜_only'] <= sel_dates[1])]
    if h_name: df_5 = df_5[df_5['상품명'].str.contains(h_name, case=False)]
    if h_time != "전체 회차": df_5 = df_5[df_5['날짜'].dt.strftime('%Y-%m-%d %H:%M:%S') == h_time]

    target_cols = ['날짜', '업체명', '상품명', '옵션', '공급처상품명', '가용재고', '기존리오더', '입고수량', '추가발주수량', '권장발주', '메모']
    # 시트에 없는 컬럼 대응
    for col in target_cols:
        if col not in df_5.columns: df_5[col] = 0

    st.dataframe(df_5[target_cols].sort_values('날짜', ascending=False), use_container_width=True, hide_index=True)

# ------------------------------------------------------------------
# 6️⃣단계: 실시간 리오더 현황판 (업체별 상황판 포함)
# ------------------------------------------------------------------
if st.session_state.get('analyzed') and 'master_log' in st.session_state:
    st.divider()
    st.header("📊 6단계: 실시간 리오더 현황판")
    m_df = st.session_state.master_log.copy()
    m_df['날짜_only'] = pd.to_datetime(m_df['날짜']).dt.date
    qty_col = '추가발주수량' if '추가발주수량' in m_df.columns else '추가발주'
    in_col = '입고수량' if '입고수량' in m_df.columns else '입고차감'
    
    f1, f2, f3 = st.columns(3)
    with f1: r_date = st.date_input("📅 기준 날짜 범위", [m_df['날짜_only'].min(), m_df['날짜_only'].max()], key="r_date")
    with f2: r_name = st.text_input("🔍 상품명 검색", key="r_name_6")
    with f3: r_vendor = st.selectbox("🏭 업체별 보기", ["전체 업체"] + list(m_df['업체명'].unique()), key="r_vendor")

    df_6 = m_df.copy()
    if len(r_date) == 2: df_6 = df_6[(df_6['날짜_only'] >= r_date[0]) & (df_6['날짜_only'] <= r_date[1])]
    if r_name: df_6 = df_6[df_6['상품명'].str.contains(r_name, case=False)]
    if r_vendor != "전체 업체": df_6 = df_6[df_6['업체명'] == r_vendor]

    # 잔량 계산 (발주 - 입고)
    df_6['잔량'] = pd.to_numeric(df_6[qty_col], errors='coerce').fillna(0) - pd.to_numeric(df_6[in_col], errors='coerce').fillna(0)
    summary = df_6.groupby(['업체명', '상품명', '옵션']).agg({'잔량': 'sum', '메모': 'last'}).reset_index()
    summary = summary[summary['잔량'] > 0].sort_values('잔량', ascending=False)

    st.subheader("🏭 업체별 미입고 요약")
    v_sum = summary.groupby('업체명')['잔량'].sum().reset_index()
    v_cols = st.columns(4)
    for i, row in v_sum.iterrows():
        with v_cols[i % 4]: st.metric(row['업체명'], f"{int(row['잔량'])} 개")

    st.subheader("📦 품목별 상세 잔량")
    st.dataframe(summary, use_container_width=True, hide_index=True)
