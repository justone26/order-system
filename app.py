import streamlit as st
import pandas as pd
import numpy as np
import re
import unicodedata
import gspread
from datetime import datetime, timedelta, timezone
import streamlit.components.v1 as components

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
        return client.open_by_key("1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg")
    except Exception as e:
        st.error(f"📡 시트 연결 실패: {e}")
        return None

def load_reorder_data():
    sh = get_sheet()
    r_map = {}
    if sh:
        try:
            ws = sh.worksheet("발주기록")
            logs = ws.get_all_values()
            if len(logs) > 1:
                df_l = pd.DataFrame(logs[1:], columns=[c.strip() for c in logs[0]])
                df_l['k'] = df_l.apply(lambda r: super_clean(r.iloc[1]) + super_clean(r.iloc[2]), axis=1)
                r_map = df_l.groupby('k').apply(lambda x: x.iloc[:, 6].apply(to_i).sum()).to_dict()
        except: pass
    return r_map

# --- [메인 로직] ---

# 1️⃣단계: 파일 업로드
st.header("1️⃣ 파일 업로드 및 데이터 로드")
up_file = st.file_uploader("엑셀 파일을 업로드하세요.", type=['xlsx', 'xls'])

if st.button("🔄 현재 화면 데이터 초기화", use_container_width=True):
    st.session_state.clear()
    st.rerun()

if up_file:
    if 'df_raw' not in st.session_state:
        st.session_state.df_raw = pd.read_excel(up_file)
        st.session_state.r_map = load_reorder_data()
        st.session_state.analyzed = False
        st.success("✅ 파일 로드 및 리오더 값 동기화 완료!")

# 2️⃣~3️⃣단계: 매핑 및 분석 실행
if 'df_raw' in st.session_state:
    st.divider()
    df_work = st.session_state.df_raw
    cols = df_work.columns.tolist()
    st.subheader("⚙️ 2️⃣단계: 매핑 설정")
    
    # ... (생략: 기존 매핑 selectbox 코드들은 동일) ...
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
            'av': avail, 't3': t3d, 't7': t1w, 'lt': lt, 'ss': ss
        }

        df = st.session_state.df_raw.copy()
        r_map = load_reorder_data()
        today = datetime.now(KST).date()

        # 데이터 계산
        df[avail] = pd.to_numeric(df[avail], errors='coerce').fillna(0).astype(int)
        df[t3d] = pd.to_numeric(df[t3d], errors='coerce').fillna(0).astype(int)
        df['clean_k'] = df.apply(lambda r: super_clean(r[item]) + super_clean(r[option]), axis=1)
        
        # ⭐ 이름 중요: '기존리오더', '일판매량', '권장발주수량'
        df['기존리오더'] = df['clean_k'].map(r_map).fillna(0).astype(int).clip(lower=0)
        
        def get_daily_avg(row):
            try:
                r_dt = pd.to_datetime(row[reg_date]).date()
                days = max(1, min((today - r_dt).days, 7))
                # 💡 소수점 반올림하여 정수로 저장
                return int(round(to_i(row[t1w]) / days, 0))
            except: 
                return int(round(to_i(row[t1w]) / 7, 0))

        df['일판매량'] = df.apply(get_daily_avg, axis=1)
        df['권장발주수량'] = ((df['일판매량'] * (lt + ss)) - (df[avail] + df['기존리오더'])).clip(lower=0).astype(int)
        
        def status_check(row):
            if "품절" in str(row[sold_out]): return "🚫 품절"
            return "🚨 발주필요" if row['권장발주수량'] > 0 else "✅ 정상"
        df['상태'] = df.apply(status_check, axis=1)
        
        # 필수 컬럼 생성
        df['입고차감'] = 0
        df['추가발주'] = 0
        df['비고(메모)'] = ""
        
        st.session_state.df_final = df
        st.session_state.analyzed = True
        st.rerun()


# ------------------------------------------------------------------
# [통합 4단계: 실시간 재고 편집 및 최종 발주 확정]
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    if 'p' not in st.session_state:
        st.warning("⚠️ 2단계에서 [분석 실행]을 먼저 눌러주세요.")
        st.stop()

    st.divider()
    st.header("📊 4단계: 입고 관리 및 최종 발주 확정")

    # 2단계에서 저장한 매핑 정보 불러오기
    p = st.session_state.p
    vnd_c, itm_c, opt_c, vit_c, avl_c, t3_c = p['vn'], p['it'], p['op'], p['vi'], p['av'], p['t3']

    # 원본 데이터 복사
    df_disp = st.session_state.df_final.copy()
    
    # [데이터 전처리] 숫자 타입 강제 변환 (에러 방지)
    num_fields = [avl_c, '기존리오더', '입고차감', '추가발주', t3_c, '일판매량', '권장발주수량']
    for col in num_fields:
        if col in df_disp.columns:
            df_disp[col] = pd.to_numeric(df_disp[col], errors='coerce').fillna(0).astype(int)

    # 1. 필터 UI 설정
    f1, f2 = st.columns([1, 2])
    with f1: 
        f_mode = st.selectbox("🚦 상태 필터", ["전체보기", "🚨 발주필요(세트)", "✅ 정상", "🚫 품절"], index=1)
    with f2: 
        s_query = st.text_input("🔍 검색 (상품명/옵션)")

    # 2. 필터 로직 적용 (품절 제외 및 세트 보기)
    if f_mode == "🚨 발주필요(세트)":
        # 상태가 '품절'이 아닌 것들 중에서 발주가 필요한 상품명(세트) 추출
        df_active = df_disp[df_disp['상태'] != "🚫 품절"]
        need_items = df_active[df_active['권장발주수량'] > 0][itm_c].unique()
        df_disp = df_active[df_active[itm_c].isin(need_items)]
    elif f_mode == "✅ 정상":
        df_disp = df_disp[df_disp['상태'] == "✅ 정상"]
    elif f_mode == "🚫 품절":
        df_disp = df_disp[df_disp['상태'] == "🚫 품절"]

    # 검색어 필터
    if s_query:
        df_disp = df_disp[df_disp[itm_c].astype(str).str.contains(s_query, case=False) | 
                          df_disp[opt_c].astype(str).str.contains(s_query, case=False)]

    # 3. 컬럼 순서 설정 및 존재 여부 체크 (KeyError 방지)
    # [순서]: 상태 → 공급처 → 상품명 → 옵션 → 공급처상품명 → 가용재고 → 기존리오더 → 입고차감 → 추가발주 → 3일판매 → 일판매량 → 권장발주수량 → 비고(메모)
    disp_cols = [
        '상태', vnd_c, itm_c, opt_c, vit_c, avl_c, 
        '기존리오더', '입고차감', '추가발주', t3_c, 
        '일판매량', '권장발주수량', '비고(메모)'
    ]
    
    for c in disp_cols:
        if c not in df_disp.columns:
            df_disp[c] = "" if "메모" in c else 0

    # 4. 통합 에디터 실행
    with st.form("v4_final_integrated_form"):
        st.info("💡 '입고차감', '추가발주', '비고(메모)' 컬럼만 수정 가능합니다. 기존잔량은 수정되지 않습니다.")
        
        edited_df = st.data_editor(
            df_disp[disp_cols],
            use_container_width=True,
            hide_index=True,
            column_config={
                '상태': st.column_config.TextColumn("상태", disabled=True),
                vnd_c: st.column_config.TextColumn("공급처", disabled=True),
                itm_c: st.column_config.TextColumn("상품명", disabled=True, width="medium"),
                opt_c: st.column_config.TextColumn("옵션", disabled=True),
                vit_c: st.column_config.TextColumn("공급처상품명", disabled=True),
                avl_c: st.column_config.NumberColumn("가용재고", disabled=True, format="%d"),
                
                # 수정 불가 컬럼 (보호)
                "기존리오더": st.column_config.NumberColumn("📦 기존잔량", disabled=True, format="%d"),
                t3_c: st.column_config.NumberColumn("3일판매", disabled=True, format="%d"),
                "일판매량": st.column_config.NumberColumn("평균판매", disabled=True, format="%d"),
                "권장발주수량": st.column_config.NumberColumn("💡 권장", disabled=True, format="%d"),
                
                # 수정 가능 컬럼
                "입고차감": st.column_config.NumberColumn("📥 입고(-)", min_value=0, format="%d"),
                "추가발주": st.column_config.NumberColumn("➕ 발주(+)", min_value=0, format="%d"),
                "비고(메모)": st.column_config.TextColumn("비고(메모)", width="medium")
            },
            key="final_stable_editor_v8"
        )
        
        btn_save = st.form_submit_button("💾 데이터 최종 저장 및 시트 전송", use_container_width=True, type="primary")

    # 5. 저장 로직
    if btn_save:
        # 입고차감이나 추가발주 값이 있는 행만 추출
        change_list = edited_df[(edited_df['입고차감'] > 0) | (edited_df['추가발주'] > 0)]
        
        if not change_list.empty:
            try:
                sh = get_sheet()
                ws_log = sh.worksheet("발주기록")
                now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
                rows = []
                
                for _, r in change_list.iterrows():
                    # 입고는 빼고, 발주는 더해서 순 발주량 계산
                    net_qty = int(r['추가발주']) - int(r['입고차감'])
                    m_txt = str(r['비고(메모)'])
                    
                    # 메모 자동 보충 (입고만 기록할 경우)
                    if r['입고차감'] > 0 and r['추가발주'] == 0:
                        m_txt = f"[입고차감] {m_txt}".strip()
                    
                    rows.append([
                        now_s,          # 일시
                        str(r[itm_c]),  # 상품명
                        str(r[opt_c]),  # 옵션
                        str(r[vit_c]),  # 공급처상품명
                        int(r[avl_c]),  # 가용재고
                        int(r['기존리오더']), # 기존잔량
                        net_qty,        # 수량
                        int(r['권장발주수량']), # 권장수량
                        m_txt,          # 메모
                        str(r[vnd_c])   # 공급처
                    ])
                
                ws_log.append_rows(rows)
                st.success(f"✅ 총 {len(rows)}건의 기록이 '발주기록' 시트에 전송되었습니다!")
                
                # 캐시 삭제 및 화면 새로고침
                st.cache_data.clear()
                time.sleep(1.5)
                st.rerun()
                
            except Exception as e:
                st.error(f"❌ 구글 시트 저장 중 오류가 발생했습니다: {e}")
        else:
            st.warning("⚠️ 입력된 입고/발주 수량이 없습니다.")

        
                
    # --- 6️⃣단계: 전체 히스토리 관리 ---
        st.divider()
        st.header("📜 6단계: 전체 히스토리 관리")
        
        h1, h2, h3, h4 = st.columns([1.2, 0.8, 1.5, 1.5])
        with h1:
            today = datetime.now(KST).date()
            d_range = st.date_input("🗓️ 1. 조회 범위", value=(today, today), key="v6_dr")
        with h2:
            st.write(""); st.write("")
            search_trigger = st.button("🔍 2. 내역 조회", use_container_width=True, type="primary")

        if 'v6_data' not in st.session_state: st.session_state.v6_data = None
        if 'v6_sessions' not in st.session_state: st.session_state.v6_sessions = []

        if search_trigger:
            sh = get_sheet()
            if sh:
                all_h = sh.worksheet("발주기록").get_all_values()
                if len(all_h) > 1:
                    df_all = pd.DataFrame(all_h[1:], columns=["발주시간", "상품명", "옵션", "공급처상품명", "가용재고", "리오더잔량", "추가발주", "발주권장", "메모", "업체명"])
                    df_all["날짜_만"] = df_all["발주시간"].str.slice(0, 10)
                    s_d = d_range[0].strftime('%Y-%m-%d'); e_d = d_range[1].strftime('%Y-%m-%d') if len(d_range)>1 else s_d
                    df_filt = df_all[(df_all["날짜_만"] >= s_d) & (df_all["날짜_만"] <= e_d)].copy()
                    st.session_state.v6_data = df_filt
                    st.session_state.v6_sessions = sorted(df_filt["발주시간"].unique(), reverse=True)
                else: st.info("저장된 내역이 없습니다.")

        with h3: h_q = st.text_input("🔍 3. 상품명/옵션 검색", key="v6_q")
        with h4:
            if st.session_state.v6_sessions:
                s_opts = ["📊 선택 범위 전체 합산"] + [f"{len(st.session_state.v6_sessions)-i}회차 ({t[5:16]})" for i, t in enumerate(st.session_state.v6_sessions)]
                sel_label = st.selectbox("📦 4. 회차 선택", s_opts)
            else: st.selectbox("📦 4. 회차 선택", ["조회 결과 없음"], disabled=True); sel_label = None

        if st.session_state.v6_data is not None and sel_label:
            df_disp = st.session_state.v6_data.copy()
            for c in ["가용재고", "리오더잔량", "추가발주", "발주권장"]: df_disp[c] = pd.to_numeric(df_disp[c], errors='coerce').fillna(0)
            
            if sel_label == "📊 선택 범위 전체 합산":
                df_disp = df_disp.groupby(["업체명", "상품명", "옵션", "공급처상품명"], as_index=False).agg({"발주시간":"max", "가용재고":"last", "리오더잔량":"last", "추가발주":"sum", "발주권장":"last", "메모": lambda x: " / ".join(set(filter(None, x.astype(str))))})
            else:
                t_time = st.session_state.v6_sessions[s_opts.index(sel_label)-1]
                df_disp = df_disp[df_disp["발주시간"] == t_time].copy()

            if h_q: df_disp = df_disp[df_disp["상품명"].str.contains(h_q, case=False) | df_disp["옵션"].str.contains(h_q, case=False)]
            st.dataframe(df_disp[["발주시간", "업체명", "상품명", "옵션", "공급처상품명", "가용재고", "리오더잔량", "추가발주", "발주권장", "메모"]], use_container_width=True, hide_index=True)
            


# ------------------------------------------------------------------
# [6단계: 실시간 리오더 현황 - 구글 시트 연동]
# ------------------------------------------------------------------
# 에러 방지를 위한 변수 사전 선언 (NameError 해결)
KST = timezone(timedelta(hours=9))
now_now = datetime.now(KST)
today_val = now_now.date()
default_start = (now_now - timedelta(days=30)).date()

st.divider()
st.header("📊 6단계: 실시간 리오더 현황")

with st.expander("📝 시트 데이터 기반 현재 리오더 잔량 확인", expanded=False):
    try:
        f1, f2 = st.columns([2, 1])
        with f1:
            # NameError 방지를 위해 사전에 정의한 변수 사용
            dr = st.date_input(
                "🗓️ 조회 기간", 
                value=(default_start, today_val), 
                key="v6_dr"
            )
        
        if len(dr) == 2:
            start_d, end_d = dr
            sh = get_sheet()
            ws_log = sh.worksheet("발주기록")
            log_data = ws_log.get_all_records()

            if log_data:
                df_l = pd.DataFrame(log_data)
                df_l['일시'] = pd.to_datetime(df_l['일시']).dt.date
                
                # 기간 필터링
                mask = (df_l['일시'] >= start_d) & (df_l['일시'] <= end_d)
                df_filtered = df_l.loc[mask]

                if not df_filtered.empty:
                    # 상품명+옵션별로 수량 합계 계산 (실시간 잔량)
                    df_res = df_filtered.groupby(['상품명', '옵션'])['수량'].sum().reset_index()
                    df_res.columns = ['상품명', '옵션', '현재리오더잔량']
                    
                    # 잔량이 0보다 큰 것만 표시
                    df_res = df_res[df_res['현재리오더잔량'] > 0].sort_values(by='현재리오더잔량', ascending=False)
                    
                    st.dataframe(
                        df_res,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "현재리오더잔량": st.column_config.NumberColumn("미입고 잔량", format="%d 📦")
                        }
                    )
                else:
                    st.info("선택한 기간 내에 발주 기록이 없습니다.")
            else:
                st.info("데이터가 없습니다.")
                
    except Exception as e:
        st.error(f"❌ 6단계 오류 발생: {e}")
