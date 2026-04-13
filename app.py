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

# 2️⃣단계: 매핑 설정
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

    # 3️⃣단계: 분석 설정
    st.divider()
    st.subheader("⚙️ 3️⃣단계: 분석 설정")
    clt, css = st.columns(2)
    lt, ss = clt.number_input("리드타임 (일)", value=10), css.number_input("안전재고 (일 수)", value=7)

    if st.button("🚀 분석 실행 / 업데이트", type="primary", use_container_width=True):
        # ⭐ [핵심 추가] 4단계에서 AttributeError를 방지하기 위해 매핑 정보를 p에 저장
        st.session_state.p = {
            'so': sold_out, 'it': item, 'op': option, 'vn': vendor, 'vi': v_item_col,
            'av': avail, 't3': t3d, 't7': t1w, 'lt': lt, 'ss': ss
        }

        df = st.session_state.df_raw.copy()
        r_map = load_reorder_data()
        today = datetime.now(KST).date()

        # 숫자 변환 안전하게 처리
        df[avail] = pd.to_numeric(df[avail], errors='coerce').fillna(0).astype(int)

        def get_daily_avg(row):
            try:
                r_dt = pd.to_datetime(row[reg_date]).date()
                diff = (today - r_dt).days
                days = max(1, min(diff, 7)) 
                return int(round(to_i(row[t1w]) / days))
            except: 
                return int(round(to_i(row[t1w]) / 7))

        df['일판매량'] = df.apply(get_daily_avg, axis=1)
        df['clean_k'] = df.apply(lambda r: super_clean(r[item]) + super_clean(r[option]), axis=1)
        df['기존리오더'] = df['clean_k'].map(r_map).fillna(0).astype(int).clip(lower=0)
        
        # 권장발주수량 계산
        df['권장발주수량'] = ((df['일판매량'] * (lt + ss)) - (df[avail] + df['기존리오더'])).clip(lower=0).astype(int)
        
        def status_check(row):
            if "품절" in str(row[sold_out]): return "🚫 품절"
            return "🚨 긴급" if row['권장발주수량'] > 0 else "✅ 정상"
        
        df['상태'] = df.apply(status_check, axis=1)
        
        # 정렬 로직
        is_emg = df.groupby(item)['상태'].transform(lambda x: any(x == "🚨 긴급"))
        df['sort_group'] = np.where(is_emg, 0, 1)
        df = df.sort_values(by=['sort_group', item, option], ascending=[True, True, True])
        
        st.session_state.df_final = df
        st.session_state.analyzed = True
        st.rerun()


# ------------------------------------------------------------------
# [통합 4단계: 입고 관리 및 최종 발주 확정]
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    # 🚨 AttributeError 방지용 안전장치
    if 'p' not in st.session_state:
        st.warning("⚠️ 매핑 정보(p)가 세션에 없습니다. 2단계에서 [분석 실행] 버튼을 다시 눌러주세요.")
        st.stop()

    st.divider()
    st.header("📊 4단계: 입고 관리 및 최종 발주 확정")
    st.info("💡 **입고차감**은 들어온 수량만큼 입력, **추가발주**는 새로 주문할 수량을 입력하세요.")

    # 2단계에서 저장한 매핑 값 불러오기
    p = st.session_state.p
    s_out, item_col, opt_col = p['so'], p['it'], p['op']
    vnd_col, v_it_col = p['vn'], p['vi']
    avl_col = p['av']

    # 1. UI 필터링
    f1, f2, f3 = st.columns([1, 2, 1])
    with f1: f_mode = st.selectbox("🚦 상태 필터", ["전체보기", "🚨 발주필요", "✅ 정상", "🚫 품절"], index=1)
    with f2: s_query = st.text_input("🔍 상품명/옵션 검색")
    with f3: in_date = st.date_input("🗓️ 입고 기록 날짜", datetime.now(KST).date())

    # 2. 데이터 준비
    df_disp = st.session_state.df_final.copy()
    
    # 필터 적용 로직
    if f_mode == "🚨 발주필요": df_disp = df_disp[df_disp['상태'] == "🚨 긴급"]
    elif f_mode == "✅ 정상": df_disp = df_disp[df_disp['상태'] == "✅ 정상"]
    elif f_mode == "🚫 품절": df_disp = df_disp[df_disp['상태'] == "🚫 품절"]

    if s_query:
        df_disp = df_disp[df_disp[item_col].astype(str).str.contains(s_query, case=False) | 
                          df_disp[opt_col].astype(str).str.contains(s_query, case=False)]

    # 3. 통합 에디터 화면
    # 표시할 컬럼 정의
    disp_cols = [vnd_col, item_col, opt_col, v_it_col, avl_col, '기존리오더', '입고차감', '추가발주', '권장발주수량', '메모']
    
    edited_df = st.data_editor(
        df_disp[disp_cols],
        use_container_width=True,
        hide_index=True,
        column_config={
            vnd_col: st.column_config.TextColumn("공급처", disabled=True),
            item_col: st.column_config.TextColumn("상품명", disabled=True, width="medium"),
            opt_col: st.column_config.TextColumn("옵션", disabled=True),
            v_it_col: st.column_config.TextColumn("공급처상품명", disabled=True),
            avl_col: st.column_config.NumberColumn("가용재고", disabled=True),
            "기존리오더": st.column_config.NumberColumn("📦 리오더잔량", disabled=True),
            "입고차감": st.column_config.NumberColumn("📥 입고(-)", min_value=0, help="들어온 수량 입력"),
            "추가발주": st.column_config.NumberColumn("➕ 발주(+)", min_value=0, help="새로 주문할 수량 입력"),
            "권장발주수량": st.column_config.NumberColumn("💡 권장", disabled=True),
            "메모": st.column_config.TextColumn("📝 메모", width="medium")
        },
        key="integrated_editor"
    )

    # 4. 저장 및 다운로드 로직
    c1, c2 = st.columns(2)
    
    if c1.button("💾 데이터 최종 저장 (시트 전송)", type="primary", use_container_width=True):
        # 입고차감이나 추가발주가 있는 행만 필터링
        change_list = edited_df[(edited_df['입고차감'] > 0) | (edited_df['추가발주'] > 0)]
        
        if not change_list.empty:
            try:
                sh = get_sheet()
                ws_log = sh.worksheet("발주기록")
                now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
                rows_to_save = []
                
                for _, r in change_list.iterrows():
                    # 수량 계산: 추가발주(+) - 입고차감(-)
                    net_qty = int(r['추가발주']) - int(r['입고차감'])
                    memo_str = str(r['메모'])
                    if r['입고차감'] > 0 and r['추가발주'] == 0:
                        memo_str = f"[입고차감] {memo_str}".strip()

                    # 발주기록 시트 형식 (10개 컬럼 기준)
                    rows_to_save.append([
                        now_s, str(r[item_col]), str(r[opt_col]), str(r[v_it_col]), 
                        int(r[avl_col]), int(r['기존리오더']), net_qty, 
                        int(r['권장발주수량']), memo_str, str(r[vnd_col])
                    ])
                
                ws_log.append_rows(rows_to_save)
                st.success(f"✅ 총 {len(rows_to_save)}건 시트 반영 완료!")
                st.cache_data.clear() # 캐시 초기화하여 최신 데이터 불러오기 유도
                time.sleep(1)
                st.rerun()
            except Exception as e:
                st.error(f"❌ 저장 오류: {e}")
        else:
            st.warning("⚠️ 변경된 수량이 없습니다.")

    with c2:
        df_down = edited_df[edited_df['추가발주'] > 0].copy()
        if not df_down.empty:
            csv_data = df_down[[vnd_col, item_col, opt_col, v_it_col, '추가발주']].rename(columns={'추가발주': '수량'})
            st.download_button(
                label="📥 추가 발주서(CSV) 다운로드",
                data=csv_data.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig'),
                file_name=f"발주서_{datetime.now(KST).strftime('%m%d_%H%M')}.csv",
                mime="text/csv",
                use_container_width=True
            )
        else:
            st.button("📥 다운로드 (발주건 없음)", disabled=True, use_container_width=True)

        
        

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

        # --- 7️⃣단계: 실시간 리오더 최종 잔량 상황판 ---
        st.divider()
        st.header("🚀 7단계: 실시간 리오더 최종 잔량 상황판")
        
        @st.cache_data(ttl=600)
        def get_v7():
            try:
                data = get_sheet().worksheet("발주기록").get_all_values()
                if len(data) > 1:
                    df = pd.DataFrame(data[1:], columns=["발주시간", "상품명", "옵션", "공급처상품명", "가용재고", "기존리오더", "추가발주", "발주권장", "메모", "업체명"])
                    df["기존리오더"] = pd.to_numeric(df["기존리오더"], errors='coerce').fillna(0).astype(int)
                    df["추가발주"] = pd.to_numeric(df["추가발주"], errors='coerce').fillna(0).astype(int)
                    df["최종잔량"] = df["기존리오더"] + df["추가발주"]
                    df["날짜_순수"] = df["발주시간"].str.slice(0, 10)
                    return df
            except: return None

        df_v7 = get_v7()
        if df_v7 is not None:
            f1, f2, f3, f4 = st.columns([1.2, 0.6, 1.5, 1.2])
            with f1: d7 = st.date_input("🗓️ 기간", value=((datetime.now(KST)-timedelta(days=30)).date(), today), key="v7_dr")
            with f2:
                st.write(""); st.write("")
                if st.button("🔄 업데이트", key="v7_up"): st.cache_data.clear(); st.rerun()
            with f3: q7 = st.text_input("🔍 검색", key="v7_qs")
            with f4: v7 = st.selectbox("🏭 업체", ["전체 업체"] + sorted(df_v7["업체명"].unique().tolist()))

            df_f7 = df_v7.copy()
            if isinstance(d7, (list, tuple)) and len(d7) == 2:
                df_f7 = df_f7[(df_f7["날짜_순수"] >= d7[0].strftime('%Y-%m-%d')) & (df_f7["날짜_순수"] <= d7[1].strftime('%Y-%m-%d'))]
            if v7 != "전체 업체": df_f7 = df_f7[df_f7["업체명"] == v7]
            if q7: df_f7 = df_f7[df_f7["상품명"].str.contains(q7, case=False) | df_f7["옵션"].str.contains(q7, case=False)]

            if not df_f7.empty:
                df_res7 = df_f7.groupby(["업체명", "상품명", "옵션", "공급처상품명"], as_index=False).agg({"발주시간":"max", "최종잔량":"sum", "메모": lambda x: " / ".join(dict.fromkeys(filter(None, x.astype(str))))})
                df_res7["최종잔량"] = df_res7["최종잔량"].clip(lower=0)
                
                df_v_sum = df_res7.groupby("업체명")["최종잔량"].sum().reset_index()
                df_v_sum = df_v_sum[df_v_sum["최종잔량"] > 0].sort_values("최종잔량", ascending=False)
                
                v_cols = st.columns(4)
                for i, r in enumerate(df_v_sum.itertuples()):
                    with v_cols[i % 4]: st.metric(r.업체명, f"{int(r.최종잔량):,} 개")
                
                st.write("#### 📋 상세 리스트 (미입고)")
                st.dataframe(df_res7[df_res7["최종잔량"] > 0].sort_values("발주시간", ascending=False), use_container_width=True, hide_index=True)
