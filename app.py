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
        df = st.session_state.df_raw.copy()
        r_map = load_reorder_data()
        today = datetime.now(KST).date()

        # [수정] 일판매량 계산 후 반올림하여 정수로 변환
        def get_daily_avg(row):
            try:
                r_dt = pd.to_datetime(row[reg_date]).date()
                diff = (today - r_dt).days
                days = max(1, min(diff, 7)) 
                # 반올림 처리 (.round() 사용)
                return int(round(to_i(row[t1w]) / days))
            except: 
                return int(round(to_i(row[t1w]) / 7))

        df['일판매량'] = df.apply(get_daily_avg, axis=1)
        df['clean_k'] = df.apply(lambda r: super_clean(r[item]) + super_clean(r[option]), axis=1)
        df['기존리오더'] = df['clean_k'].map(r_map).fillna(0).astype(int)
        
        av_val = pd.to_numeric(df[avail], errors='coerce').fillna(0)
        # 반올림된 일판매량 기반 권장발주수량 계산
        df['권장발주수량'] = ((df['일판매량'] * (lt + ss)) - (av_val + df['기존리오더'])).clip(lower=0).astype(int)
        
        def status_check(row):
            if "품절" in str(row[sold_out]): return "🚫 품절"
            return "🚨 긴급" if row['권장발주수량'] > 0 else "✅ 정상"
        df['상태'] = df.apply(status_check, axis=1)
        
        is_emg = df.groupby(item)['상태'].transform(lambda x: any(x == "🚨 긴급"))
        df['sort_group'] = np.where(is_emg, 0, 1)
        df = df.sort_values(by=['sort_group', item, option], ascending=[True, True, True])
        
        df['추가발주'], df['입고차감'], df['메모'] = 0, 0, ""
        st.session_state.df_final = df
        st.session_state.analyzed = True
        st.rerun()

    # 4️⃣단계: 편집 및 저장
    if st.session_state.get('analyzed'):
        st.divider()
        st.header("4️⃣ 발주 편집 및 저장")
        df_f = st.session_state.df_final.copy()
        
        f1, f2 = st.columns([1, 2])
        sel_s = f1.selectbox("🚦 상태 필터", ["전체상품", "🚨 긴급", "✅ 정상", "🚫 품절"])
        q = f2.text_input("🔎 상품명 검색")

        disp = df_f.copy()
        if sel_s == "🚨 긴급": disp = disp[disp['sort_group'] == 0]
        elif sel_s == "✅ 정상": disp = disp[disp['상태'] == "✅ 정상"]
        elif sel_s == "🚫 품절": disp = disp[disp['상태'] == "🚫 품절"]
        if q: disp = disp[disp[item].str.contains(q, case=False, na=False)]

        # 컬럼 순서 (사장님 요청 순서 고정)
        safe = [
            '상태', vendor, item, option, v_item_col, 
            avail, '기존리오더', '입고차감', '추가발주', 
            t3d, '일판매량', '권장발주수량', '메모'
        ]
        
        try:
            edit_df = st.data_editor(
                disp[safe], 
                hide_index=True, 
                use_container_width=True, 
                key="final_v26_int",
                column_config={
                    "상태": st.column_config.TextColumn("상태", width="small"),
                    vendor: st.column_config.TextColumn("공급처"),
                    item: st.column_config.TextColumn("상품명", width="large"),
                    v_item_col: st.column_config.TextColumn("공급처상품명"),
                    "입고차감": st.column_config.NumberColumn("➖ 입고수량", step=1),
                    "추가발주": st.column_config.NumberColumn("➕ 추가발주", step=1),
                    t3d: st.column_config.NumberColumn("📅 3일판매"),
                    "일판매량": st.column_config.NumberColumn("🔥 일판매량", format="%d"), # 정수 표기
                    "메모": st.column_config.TextColumn("비고", width="medium")
                }
            )

            if st.button("💾 내역 저장 및 수치 즉시 갱신", type="primary", use_container_width=True):
                to_save = edit_df[(edit_df['추가발주'] > 0) | (edit_df['입고차감'] > 0)]
                if not to_save.empty:
                    sh = get_sheet()
                    ws = sh.worksheet("발주기록")
                    now = datetime.now(KST).strftime('%Y-%m-%d %H:%M')
                    rows = [[now, str(r[item]), str(r[option]), str(r[v_item_col]), int(to_i(r[avail])), int(r['기존리오더']), int(r['추가발주']) - int(r['입고차감']), int(r['권장발주수량']), str(r['메모']), str(r[vendor])] for _, r in to_save.iterrows()]
                    ws.append_rows(rows)
                    st.success("✅ 저장 완료! 리오더 수치를 업데이트합니다.")
                    st.session_state.analyzed = False 
                    st.rerun()
        except KeyError as e:
            st.error(f"❌ 매핑 에러: {e} 컬럼을 찾을 수 없습니다.")

    # ------------------------------------------------------------------
# 6️⃣단계: 저장 내역 상세 검색 (날짜/회차/상품명 검색 후 조회)
# ------------------------------------------------------------------
st.divider()
st.header("6️⃣ 저장 내역 상세 검색")

# [1] 상단 검색 설정 영역
c1, c2, c3 = st.columns([1, 1, 1])
with c1:
    q_date_6 = st.date_input("📅 1차: 날짜 선택", value=datetime.now(KST).date(), key="final_date_6")
with c2:
    st.write("") # 간격 맞춤
    # 이 버튼을 눌러야만 아래 내용이 나옵니다.
    btn_load_6 = st.button("🚀 데이터 조회하기", use_container_width=True, type="primary", key="final_btn_6")

# [2] 조회 버튼 클릭 시 실행
if btn_load_6 or st.session_state.get('step6_active'):
    st.session_state.step6_active = True  # 조회 상태 유지
    
    sh = get_sheet()
    if sh:
        ws = sh.worksheet("발주기록")
        raw = ws.get_all_values()
        if len(raw) > 1:
            df_6 = pd.DataFrame(raw[1:], columns=[c.strip() for c in raw[0]])
            
            # 날짜/시간 데이터 정리 (화면 미노출용)
            df_6['pure_dt'] = df_6.iloc[:, 0].str.strip()
            df_6['pure_date'] = df_6['pure_dt'].str.split(' ').str[0]
            df_6['pure_time'] = df_6['pure_dt'].str.split(' ').str[1].str[:5]
            
            target_s = q_date_6.strftime('%Y-%m-%d')
            res_date = df_6[df_6['pure_date'] == target_s].copy()

            if not res_date.empty:
                # [상단 2차 필터: 회차 및 상품명]
                sub_f1, sub_f2 = st.columns(2)
                with sub_f1:
                    times = sorted(res_date['pure_time'].dropna().unique(), reverse=True)
                    q_time = st.selectbox(f"⏰ 2차: 회차 선택 ({len(times)}회)", ["전체 보기"] + times, key="final_time_6")
                with sub_f2:
                    q_item = st.text_input("🔎 3차: 상품명 검색", key="final_search_6")

                # 필터링 적용
                display_6 = res_date.copy()
                if q_time != "전체 보기":
                    display_6 = display_6[display_6['pure_time'] == q_time]
                if q_item:
                    # 상품명 컬럼(index 1) 기준 검색
                    display_6 = display_6[display_6.iloc[:, 1].str.contains(q_item, case=False)]

                # [사장님 요청 컬럼 순서 고정]
                # 날짜 => 업체명 => 상품명 => 옵션 => 공급처상품명 => 가용재고 => 기존리오더 => 입고수량(G열) => 추가발주 => 권장발주 => 메모
                col_order = [
                    display_6.columns[0],   # 날짜
                    display_6.columns[-1],  # 업체명
                    "상품명", "옵션", "공급처상품명", "가용재고", "기존리오더", 
                    display_6.columns[6],   # 입고수량(G열)
                    "추가발주", "권장발주수량", "메모"
                ]
                
                # 실제 존재하는 컬럼만 노출 (순서 강제 지정, 보조컬럼 제외)
                final_cols = [c for c in col_order if c in display_6.columns]
                
                st.dataframe(display_6[final_cols].iloc[::-1], use_container_width=True, hide_index=True)
            else:
                st.info(f"{target_s}에 저장된 내역이 없습니다.")

# ------------------------------------------------------------------
# 7️⃣단계: 실시간 리오더 최종 잔량 상황판 (독립형 업데이트)
# ------------------------------------------------------------------
st.divider()
st.header("7️⃣ 실시간 리오더 최종 잔량 상황판")

# 상단 업데이트 버튼
if st.button("📊 실시간 현황판 업데이트", key="final_btn_7", use_container_width=True, type="secondary"):
    sh = get_sheet()
    if sh:
        ws = sh.worksheet("발주기록")
        raw = ws.get_all_values()
        if len(raw) > 1:
            df_7 = pd.DataFrame(raw[1:], columns=[c.strip() for c in raw[0]])
            
            # 수량 계산 (G열: 추가발주/차감액)
            df_7['qty_val'] = df_7.iloc[:, 6].apply(to_i)
            v_col_7 = df_7.columns[-1] # 업체명
            
            # 업체/상품/옵션별 잔량 합계 계산
            df_res_7 = df_7.groupby([v_col_7, df_7.columns[1], df_7.columns[2]], as_index=False)['qty_val'].sum()
            df_res_7 = df_res_7[df_res_7['qty_val'] > 0] # 미입고만 추출
            df_res_7.columns = ['업체명', '상품명', '옵션', '잔량']

            if not df_res_7.empty:
                # 업체별 총량 메트릭
                v_sum_7 = df_res_7.groupby('업체명')['잔량'].sum().reset_index()
                m_cols = st.columns(4)
                for idx, r in enumerate(v_sum_7.itertuples()):
                    with m_cols[idx % 4]:
                        st.metric(r.업체명, f"{int(r.잔량):,} 개")
                
                st.write("#### 📋 상세 미입고 리스트")
                st.dataframe(df_res_7.sort_values('잔량', ascending=False), use_container_width=True, hide_index=True)
            else:
                st.success("🎉 현재 모든 입고가 완료되어 미입고 잔량이 없습니다!")
