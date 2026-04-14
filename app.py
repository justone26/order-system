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

    if st.button("🚀 분석 실행 / 전체 장부 업데이트", type="primary", use_container_width=True):
        st.session_state.p = {
            'so': sold_out, 'it': item, 'op': option, 'vn': vendor, 'vi': v_item_col,
            'av': avail, 't3': t3d, 't7': t1w, 'lt': lt, 'ss': ss, 'rd': reg_date
        }

        with st.spinner("📊 구글 시트 동기화 및 데이터 분석 중..."):
            try:
                # 1. 엑셀 업로드 데이터 처리
                df = st.session_state.df_raw.copy()
                today = datetime.now(KST).date()

                # 숫자 변환
                df[avail] = pd.to_numeric(df[avail], errors='coerce').fillna(0).astype(int)
                df[t1w] = pd.to_numeric(df[t1w], errors='coerce').fillna(0).astype(int)

                # ---------------------------------------------------------
                # ⭐ [핵심 추가] 5~6단계를 위한 구글 시트 전체 장부 로드
                # ---------------------------------------------------------
                sh = get_sheet()
                ws_log = sh.worksheet("발주기록")
                df_master = pd.DataFrame(ws_log.get_all_records())
                df_master.columns = [c.strip() for c in df_master.columns]
                st.session_state.master_log = df_master # 5, 6단계용 데이터 세션 저장
                
                # 기존 리오더 매핑을 위한 리워킹 (r_map 생성)
                # '추가발주수량' 또는 '추가발주' 컬럼에서 잔량 계산
                q_col = '추가발주수량' if '추가발주수량' in df_master.columns else '추가발주'
                df_master[q_col] = pd.to_numeric(df_master[q_col], errors='coerce').fillna(0)
                r_map = df_master.groupby(['상품명', '옵션'])[q_col].sum().to_dict()
                # ---------------------------------------------------------

                # 키 생성 및 리오더 매핑
                df['clean_k'] = df.apply(lambda r: (str(r[item]) + str(r[option])).replace(" ", ""), axis=1)
                
                # 리오더 매핑 (딕셔너리 키 매칭)
                def get_reorder_val(row):
                    key = (row[item], row[option])
                    return r_map.get(key, 0)
                
                df['기존리오더'] = df.apply(get_reorder_val, axis=1).fillna(0).astype(int).clip(lower=0)
                
                # 평균 판매량 및 권장발주 계산
                def get_daily_avg(row):
                    try:
                        r_dt = pd.to_datetime(row[reg_date]).date()
                        days = max(1, min((today - r_dt).days, 7))
                        return int(round(pd.to_numeric(row[t1w], errors='coerce') / days, 0))
                    except: return int(round(pd.to_numeric(row[t1w], errors='coerce') / 7, 0))

                df['일판매량'] = df.apply(get_daily_avg, axis=1)
                df['권장발주수량'] = ((df['일판매량'] * (lt + ss)) - (df[avail] + df['기존리오더'])).clip(lower=0).astype(int)
                
                def status_check(row):
                    if "품절" in str(row[sold_out]): return "🚫 품절"
                    return "🚨 발주필요" if row['권장발주수량'] > 0 else "✅ 정상"
                
                df['상태'] = df.apply(status_check, axis=1)
                
                # 4단계 입력용 기본 컬럼 세팅
                df['입고차감'] = 0
                df['추가발주'] = 0
                df['비고(메모)'] = ""
                
                # 모든 분석 완료 후 세션 저장
                st.session_state.df_final = df
                st.session_state.analyzed = True # 4, 5, 6단계가 동시에 열리게 됨
                
                st.success("✅ 분석 및 장부 동기화가 완료되었습니다!")
                time.sleep(0.5)
                st.rerun()
                
            except Exception as e:
                st.error(f"분석 오류: {e}")

# 4️⃣단계: 입고 관리 및 최종 발주 확정
if st.session_state.get('analyzed'):
    st.divider()
    st.header("📊 4단계: 입고 관리 및 최종 발주 확정")
    
    p = st.session_state.p
    df_disp = st.session_state.df_final.copy()
    
    # 필터 로직
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
                '권장발주수량': st.column_config.NumberColumn("💡 권장", disabled=True),
                '비고(메모)': st.column_config.TextColumn("📝 메모") # 사장님이 직접 쓰실 칸
            }
        )
        btn_save = st.form_submit_button("💾 데이터 최종 저장 및 시트 전송", use_container_width=True, type="primary")

    if btn_save:
        # 변경사항(입고나 발주 수량이 있는 경우)만 추출
        change_list = edited_df[(edited_df['입고차감'] > 0) | (edited_df['추가발주'] > 0)]
        
        if not change_list.empty:
            try:
                sh = get_sheet()
                ws_log = sh.worksheet("발주기록")
                now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
                
                rows = []
                for _, r in change_list.iterrows():
                    # ⭐ [자동 메모 생성 로직]
                    memo_parts = []
                    if r['입고차감'] > 0:
                        memo_parts.append(f"-{int(r['입고차감'])} 입고")
                    if r['추가발주'] > 0:
                        memo_parts.append(f"{int(r['추가발주'])} 발주")
                    
                    # 자동 생성된 문구 (예: "-50 입고 / 100 발주")
                    auto_memo = " / ".join(memo_parts)
                    
                    # 사장님이 직접 쓴 메모 결합
                    user_memo = str(r['비고(메모)']).strip() if r['비고(메모)'] and str(r['비고(메모)']) != "None" else ""
                    
                    if user_memo:
                        final_memo = f"[{auto_memo}] {user_memo}"
                    else:
                        final_memo = auto_memo

                    # 수량 계산 (추가발주 - 입고차감)
                    net_qty = int(r['추가발주']) - int(r['입고차감'])

                    # 시트 행 데이터 구성
                    rows.append([
                        now_s,          # 날짜
                        r[p['it']],      # 상품명
                        r[p['op']],      # 옵션
                        r[p['vi']],      # 공급처상품명
                        r[p['av']],      # 가용재고
                        r['기존리오더'],  # 기존잔량
                        net_qty,        # 추가발주(수량) -> net_qty로 저장
                        r['권장발주수량'],
                        final_memo,      # ⭐ 위에서 만든 자동 메모
                        r[p['vn']]       # 업체명
                    ])
                
                ws_log.append_rows(rows)
                st.success(f"✅ {len(rows)}건 저장 완료!")
                st.cache_data.clear()
                # 5단계를 위해 로드된 데이터 초기화
                if 'master_log' in st.session_state:
                    del st.session_state.master_log
                time.sleep(1)
                st.rerun()
                
            except Exception as e: 
                st.error(f"저장 실패: {e}")

# 5️⃣단계: 전체 발주/입고 히스토리
if st.session_state.get('analyzed') and 'master_log' in st.session_state:
    st.divider()
    st.header("📜 5단계: 전체 히스토리 기록")
    
    m_df = st.session_state.master_log.copy()
    m_df['날짜'] = pd.to_datetime(m_df['날짜'], errors='coerce')
    m_df['날짜_only'] = m_df['날짜'].dt.date
    
    # 상단 검색 필터
    c1, c2, c3 = st.columns(3)
    with c1:
        sel_dates = st.date_input("📅 조회 날짜", [m_df['날짜_only'].max() - timedelta(days=7), m_df['날짜_only'].max()], key="h_date")
    with c2:
        h_name = st.text_input("🔍 상품명 검색", key="h_name")
    with c3:
        # 회차(시간대) 선택
        times = sorted(m_df['날짜'].unique(), reverse=True)
        t_options = ["전체 회차"] + [t.strftime('%Y-%m-%d %H:%M:%S') for t in times]
        h_time = st.selectbox("⏰ 회차(시간) 선택", t_options, key="h_time")

    # 필터 적용
    df_5 = m_df.copy()
    if len(sel_dates) == 2:
        df_5 = df_5[(df_5['날짜_only'] >= sel_dates[0]) & (df_5['날짜_only'] <= sel_dates[1])]
    if h_name:
        df_5 = df_5[df_5['상품명'].str.contains(h_name, case=False)]
    if h_time != "전체 회차":
        df_5 = df_5[df_5['날짜'].dt.strftime('%Y-%m-%d %H:%M:%S') == h_time]

    # 요청하신 컬럼 순서 정렬
    target_cols = ['날짜', '업체명', '상품명', '옵션', '공급처상품명', '가용재고', '기존리오더', '입고수량', '추가발주수량', '권장발주', '메모']
    st.dataframe(df_5[target_cols].sort_values('날짜', ascending=False), use_container_width=True, hide_index=True)
    

        
       # 6️⃣단계: 실시간 리오더 현황 (대시보드)
if st.session_state.get('analyzed') and 'master_log' in st.session_state:
    st.divider()
    st.header("📊 6단계: 실시간 리오더 현황판")
    
    m_df = st.session_state.master_log.copy()
    qty_col = '추가발주수량' if '추가발주수량' in m_df.columns else '추가발주'
    m_df[qty_col] = pd.to_numeric(m_df[qty_col], errors='coerce').fillna(0)

    # 상단 검색 필터
    f1, f2, f3 = st.columns(3)
    with f1:
        r_date = st.date_input("📅 기준 날짜 범위", [m_df['날짜_only'].min(), m_df['날짜_only'].max()], key="r_date")
    with f2:
        r_name = st.text_input("🔍 상품명 검색", key="r_name")
    with f3:
        v_list = ["전체 업체"] + list(m_df['업체명'].unique())
        r_vendor = st.selectbox("🏭 업체별 보기", v_list, key="r_vendor")

    # 데이터 집계 (잔량 계산)
    df_6 = m_df.copy()
    if len(r_date) == 2:
        df_6 = df_6[(df_6['날짜_only'] >= r_date[0]) & (df_6['날짜_only'] <= r_date[1])]
    if r_name:
        df_6 = df_6[df_6['상품명'].str.contains(r_name, case=False)]
    if r_vendor != "전체 업체":
        df_6 = df_6[df_6['업체명'] == r_vendor]

    # 품목별 잔량 합산
    summary = df_6.groupby(['업체명', '상품명', '옵션']).agg({
        qty_col: 'sum',
        '메모': 'last' # 가장 최근 메모
    }).reset_index()
    summary.columns = ['업체명', '상품명', '옵션', '현재잔량', '최근처리내용']
    summary = summary[summary['현재잔량'] > 0].sort_values('현재잔량', ascending=False)

    # --- [업체별 요약 상황판] ---
    st.subheader("🏭 업체별 미입고 요약")
    v_summary = summary.groupby('업체명')['현재잔량'].sum().reset_index()
    v_cols = st.columns(len(v_summary) if len(v_summary) > 0 else 1)
    for i, row in v_summary.iterrows():
        with v_cols[i % len(v_cols)]:
            st.metric(label=row['업체명'], value=f"{int(row['현재잔량'])} 개")

    # --- [상세 현황 리스트] ---
    st.subheader("📦 품목별 상세 잔량")
    st.dataframe(summary, use_container_width=True, hide_index=True,
                 column_config={"현재잔량": st.column_config.NumberColumn("잔량", format="%d 📦")})
