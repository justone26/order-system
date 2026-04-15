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
# 1️⃣단계: 파일 업로드 및 데이터 로드 (계산 로직: 통장 잔액 방식 적용)
# ------------------------------------------------------------------
st.header("1️⃣ 파일 업로드 및 데이터 로드")

if 'uploader_key' not in st.session_state:
    st.session_state.uploader_key = 0

up_file = st.file_uploader(
    "엑셀 파일을 업로드하세요.", 
    type=['xlsx', 'xls'],
    key=f"file_uploader_{st.session_state.uploader_key}"
)

if st.button("🔄 현재 화면 데이터 초기화", use_container_width=True):
    for key in list(st.session_state.keys()):
        if key != 'uploader_key':
            del st.session_state[key]
    st.session_state.uploader_key += 1
    st.success("✅ 파일과 모든 데이터가 초기화되었습니다.")
    time.sleep(0.5)
    st.rerun()

if up_file is not None:
    if 'df_raw' not in st.session_state:
        try:
            temp_df = pd.read_excel(up_file)
            st.session_state.df_raw = temp_df
            st.session_state.analyzed = False
            
            # 🚨 [수정 핵심] 장부를 날짜순으로 정렬한 뒤 '진짜 마지막' 행을 가져옵니다.
            with st.spinner("🔄 구글 시트에서 최신 기존리오더(장부 잔액) 동기화 중..."):
                sh = get_sheet()
                ws = sh.worksheet("발주기록")
                raw_records = ws.get_all_values()
                
                if len(raw_records) > 1:
                    df_rec = pd.DataFrame(raw_records[1:], columns=[h.strip() for h in raw_records[0]])
                    def c_func(t): return "".join(str(t).split()).upper()
                    
                    df_rec['key'] = df_rec['상품명'].apply(c_func) + df_rec['옵션'].apply(c_func)
                    
                    # 1. 숫자 변환 (콤마 제거)
                    df_rec['기존리오더'] = pd.to_numeric(df_rec['기존리오더'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                    
                    # 2. 🚨 날짜/시간 정렬 (이게 없으면 예전 100장 기록을 가져올 수 있습니다)
                    # '날짜' 컬럼을 기준으로 오름차순 정렬하여 최신 기록이 아래로 가게 만듭니다.
                    df_rec['날짜_dt'] = pd.to_datetime(df_rec['날짜'], errors='coerce', format='mixed')
                    df_rec = df_rec.sort_values('날짜_dt', ascending=True)
                    
                    # 3. 마지막 행(최신 잔액 50장 시점) 추출
                    last_balance_map = df_rec.drop_duplicates('key', keep='last').set_index('key')['기존리오더'].to_dict()
                    st.session_state.reorder_ans = last_balance_map
                else:
                    st.session_state.reorder_ans = {}
                    
            st.success("✅ 엑셀 및 장부 잔량 동기화 완료!")
            st.rerun()
        except Exception as e:
            st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")
            
            
# ------------------------------------------------------------------
# 2️⃣단계: 매핑 설정 (변수 유실 방지 처리)
# ------------------------------------------------------------------
if 'df_raw' in st.session_state:
    st.divider()
    df_work = st.session_state.df_raw
    cols = df_work.columns.tolist()
    
    st.subheader("⚙️ 2️⃣단계: 매핑 설정")
    c1, c2 = st.columns(2)
    with c1:
        # 🚨 key를 지정하여 세션에 자동 저장되도록 보강
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


# ------------------------------------------------------------------
# 3️⃣단계: 분석 설정 및 실행 (실시간 장부 동기화 적용)
# ------------------------------------------------------------------
if 'df_raw' in st.session_state:
    st.divider()
    st.subheader("⚙️ 3️⃣단계: 분석 설정 및 실행")

    clt, css = st.columns(2)
    with clt: lt = st.number_input("리드타임 (일)", value=10, key="input_lt")
    with css: ss = st.number_input("안전재고 (일 수)", value=7, key="input_ss")

    if st.button("🚀 분석 실행", type="primary", use_container_width=True):
        try:
            # 🚨 [추가] 분석 버튼 누를 때마다 구글 시트에서 최신 잔액을 새로 가져옴
            with st.spinner("🔄 장부에서 최신 잔액(기존리오더) 동기화 중..."):
                sh = get_sheet()
                ws = sh.worksheet("발주기록")
                raw_records = ws.get_all_values()
                
                if len(raw_records) > 1:
                    df_rec = pd.DataFrame(raw_records[1:], columns=[h.strip() for h in raw_records[0]])
                    def c_func(t): return "".join(str(t).split()).upper()
                    
                    # 데이터 정렬 및 최신 잔액 추출 (통장 방식)
                    df_rec['key'] = df_rec['상품명'].apply(c_func) + df_rec['옵션'].apply(c_func)
                    df_rec['기존리오더'] = pd.to_numeric(df_rec['기존리오더'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                    df_rec['날짜_dt'] = pd.to_datetime(df_rec['날짜'], errors='coerce', format='mixed')
                    
                    # 날짜순 정렬 후 마지막 행(최신 잔액) 맵 생성
                    last_balance_map = df_rec.sort_values('날짜_dt').drop_duplicates('key', keep='last').set_index('key')['기존리오더'].to_dict()
                    st.session_state.reorder_ans = last_balance_map
                else:
                    st.session_state.reorder_ans = {}

            # 기존 분석 로직 시작
            p_map = {
                'so': st.session_state.get('sel_so'), 'it': st.session_state.get('sel_it'),
                'op': st.session_state.get('sel_op'), 'vn': st.session_state.get('sel_vn'),
                'vi': st.session_state.get('sel_vi'), 'av': st.session_state.get('sel_av'),
                't3': st.session_state.get('sel_t3'), 't7': st.session_state.get('sel_t7'),
                'rd': st.session_state.get('sel_rd'), 'lt': lt, 'ss': ss
            }
            st.session_state.p = p_map

            df = st.session_state.df_raw.copy()
            def c_func_final(t): return "".join(str(t).split()).upper()
            df['key'] = df[p_map['it']].apply(c_func_final) + df[p_map['op']].apply(c_func_final)

            # 위에서 새로 긁어온 ans_map을 적용 (100장이 아니라 50장으로 업데이트됨)
            ans_map = st.session_state.get('reorder_ans', {})
            df['기존리오더'] = df['key'].map(ans_map).fillna(0).astype(int)

            # 판매량 및 가용재고 계산
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

            # 권장발주수량 계산
            df['권장발주수량'] = ((df['일판매량'] * (lt + ss)) - (df[c_av] + df['기존리오더'])).clip(lower=0).astype(int)
            
            # 상태 판별 및 UI 초기화
            df['상태'] = df.apply(lambda r: "🚫 품절" if "품절" in str(r[p_map['so']]) else ("🚨 발주필요" if r['권장발주수량'] > 0 else "✅ 정상"), axis=1)
            df['입고차감'] = 0 ; df['추가발주'] = 0 ; df['비고(처리내역)'] = ""

            if 'key' in df.columns: df = df.drop(columns=['key'])
            st.session_state.df_final = df
            st.session_state.analyzed = True
            
            st.cache_data.clear()
            st.success(f"✅ 분석 완료 (최신 장고 반영 완료 / LT:{lt}일/SS:{ss}일)")
            st.rerun()

        except Exception as e:
            st.error(f"⚠️ 분석 오류: {e}")
            

# ------------------------------------------------------------------
# 4️⃣단계: 입고 관리 및 최종 저장 (통장 잔액 저장 방식 적용)
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.header("📊 4단계: 입고 관리 및 최종 발주 확정")
    
    p = st.session_state.p
    
    if 'df_final' not in st.session_state:
        st.error("데이터가 없습니다. 이전 단계를 먼저 진행해 주세요.")
        st.stop()

    # 원본 복사
    df_all = st.session_state.df_final.copy()

    # 🚨 필수 컬럼 보장
    for col in ['기존리오더', '입고차감', '추가발주', '일판매량', '권장발주수량', '비고(처리내역)', '상태']:
        if col not in df_all.columns:
            df_all[col] = 0 if '수량' in col or '리오더' in col or '차감' in col or '발주' in col else ""

    # [기능유지] 사장님표 상태 필터 및 검색 UI
    f1, f2 = st.columns([1, 2])
    with f1: 
        f_mode = st.selectbox("🚦 상태 필터", ["전체보기", "🚨 발주필요(세트)", "✅ 정상", "🚫 품절"], index=1)
    with f2: 
        s_query = st.text_input("🔍 검색 (상품명/옵션)")

    # 필터링 로직 (유지)
    df_temp = df_all.copy()
    if f_mode == "🚨 발주필요(세트)":
        need_items = df_temp[(df_temp['상태'] != "🚫 품절") & (df_temp['권장발주수량'] > 0)][p['it']].unique()
        df_temp = df_temp[df_temp[p['it']].isin(need_items)]
    elif f_mode != "전체보기":
        df_temp = df_temp[df_temp['상태'] == f_mode]
        
    if s_query:
        df_temp = df_temp[df_temp[p['it']].str.contains(s_query, case=False, na=False) | 
                           df_temp[p['op']].str.contains(s_query, case=False, na=False)]

    # 노출 컬럼 설정
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
                '상태': st.column_config.TextColumn("상태", disabled=True),
                p['vn']: st.column_config.TextColumn("공급처", disabled=True),
                p['it']: st.column_config.TextColumn("상품명", disabled=True),
                p['op']: st.column_config.TextColumn("옵션", disabled=True),
                '기존리오더': st.column_config.NumberColumn("기존리오더", disabled=True, format="%d"),
                '입고차감': st.column_config.NumberColumn("📥 입고(-)", min_value=0), 
                '추가발주': st.column_config.NumberColumn("➕ 발주(+)", min_value=0),
                '권장발주수량': st.column_config.NumberColumn("권장수량", disabled=True, format="%d"),
            }
        )
        btn_save = st.form_submit_button("🚀 최종 데이터 저장 및 시트 전송", use_container_width=True, type="primary")

    if btn_save:
        # 변경 내역 확인 (수량 변화가 있거나 메모가 있는 경우 포함)
        changed_rows = edited_df[(edited_df['입고차감'] > 0) | (edited_df['추가발주'] > 0) | (edited_df['비고(처리내역)'].str.strip() != "")].copy()
        
        if not changed_rows.empty:
            with st.spinner("🚀 통장 잔액 방식으로 장부 업데이트 중..."):
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
                        
                        # 🚨 [계산 로직 핵심] 통장 잔액 방식 적용
                        # 최종 잔액 = 현재 기존리오더 + 이번 발주량 - 이번 입고량
                        final_balance = int(r['기존리오더']) + q_val - i_val
                        
                        m_parts = []
                        if q_val > 0: m_parts.append(f"{time_short} {q_val}발주")
                        if i_val > 0: m_parts.append(f"{time_short} {i_val}입고")
                        final_memo = f"[{' '.join(m_parts)}] {user_memo}".strip()
                        
                        # 발주기록(6단계 원본)에는 계산된 'final_balance'를 [기존리오더] 컬럼에 저장
                        rows_qty.append([now_s, r[p['vn']], r[p['it']], r[p['op']], r[p['vi']], final_balance, q_val, i_val, final_memo])
                        # 히스토리에는 당시의 상황을 모두 기록
                        rows_hist.append([now_s, r[p['vn']], r[p['it']], r[p['op']], r[p['vi']], r[p['av']], r['기존리오더'], i_val, q_val, r['권장발주수량'], final_memo])

                        # 세션 동기화
                        mask = (st.session_state.df_final[p['it']] == r[p['it']]) & (st.session_state.df_final[p['op']] == r[p['op']])
                        st.session_state.df_final.loc[mask, '기존리오더'] = final_balance
                        st.session_state.df_final.loc[mask, '입고차감'] = 0
                        st.session_state.df_final.loc[mask, '추가발주'] = 0

                    if rows_qty: ws_qty.append_rows(rows_qty, value_input_option='USER_ENTERED')
                    if rows_hist: ws_hist.append_rows(rows_hist, value_input_option='USER_ENTERED')
                    
                    st.cache_data.clear()
                    st.success(f"✅ 저장 완료! 최신 잔액({final_balance}개 등)이 장부에 기록되었습니다.")
                    time.sleep(1)
                    st.rerun()
                except Exception as e:
                    st.error(f"저장 중 오류: {e}")
        else:
            st.warning("⚠️ 변경 내역이 없습니다.")


# ------------------------------------------------------------------
# 6️⃣단계: 리오더 현황판 (TOP 3 상품명 통합 및 중복 제거 버전)
# ------------------------------------------------------------------
def render_step6():
    if not (st.session_state.get('analyzed') or st.session_state.get('show_step6')):
        return

    st.markdown("---")
    st.markdown("### 📈 6단계: 실시간 리오더 현황판 (최신 잔액 기준)")
    
    try:
        sh = get_sheet()
        ws_qty = sh.worksheet("발주기록")
        data = ws_qty.get_all_values()
        if len(data) <= 1: return
        
        df_log = pd.DataFrame(data[1:], columns=[h.strip() for h in data[0]])
        
        for col in ['기존리오더', '추가발주', '입고수량']:
            if col in df_log.columns:
                df_log[col] = pd.to_numeric(df_log[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        
        df_log['날짜_dt'] = pd.to_datetime(df_log['날짜'], errors='coerce', format='mixed')
    except Exception as e:
        st.error(f"데이터 로드 중 오류: {e}"); return

    # [1] UI 레이아웃
    c1, c2, c3 = st.columns([1, 2, 1.5])
    with c1:
        st.markdown("<p style='margin-bottom: 8px; font-size: 14px; font-weight: normal;'>🔄 데이터 갱신</p>", unsafe_allow_html=True)
        if st.button("최신 자료 업데이트", use_container_width=True, key="btn_update_compact"):
            st.cache_data.clear()
            st.rerun()
    with c2:
        sel_s = st.text_input("🔍 통합 상품명 검색", placeholder="상품명을 입력하세요", key="s6_search_compact")
    with c3:
        v_list = ["전체 공급처"] + sorted(df_log['공급처'].unique().tolist())
        sel_v = st.selectbox("🏭 공급처 필터", v_list, key="s6_vendor_compact")

    # [2] 데이터 그룹화 (통장 잔액 방식)
    def c_func(t): return "".join(str(t).split()).upper()
    df_log['key'] = df_log['상품명'].apply(c_func) + df_log['옵션'].apply(c_func)
    
    # 최신 잔액행만 추출
    grouped = df_log.sort_values('날짜_dt').drop_duplicates('key', keep='last').copy()
    
    memo_map = df_log.groupby('key')['메모'].apply(
        lambda x: " / ".join([str(i).strip() for i in x.tail(5) if str(i).strip() not in ["", "None", "nan"]])
    ).to_dict()
    
    grouped['최종잔량'] = grouped['기존리오더']
    grouped['최종메모'] = grouped['key'].map(memo_map)

    # 필터링
    filtered_grouped = grouped[grouped['최종잔량'] > 0].copy()
    if sel_s: filtered_grouped = filtered_grouped[filtered_grouped['상품명'].str.contains(sel_s, case=False)]
    if sel_v != "전체 공급처": filtered_grouped = filtered_grouped[filtered_grouped['공급처'] == sel_v]

    # [3] 🏢 업체별 요약 (🚨 TOP 3 상품명 통합 로직 적용)
    st.markdown("#### 🏢 업체별 미입고 및 주요 상품")
    v_sum = filtered_grouped.groupby('공급처')['최종잔량'].sum().reset_index()
    v_sum = v_sum.sort_values('최종잔량', ascending=False)
    
    if not v_sum.empty:
        v_cols = st.columns(min(len(v_sum), 4))
        for i, (idx, row) in enumerate(v_sum.iterrows()):
            if i < 4:
                v_name = row['공급처']
                with v_cols[i]:
                    st.metric(v_name, f"{int(row['최종잔량'])}개 잔량")
                    
                    # 🚨 [핵심 수정] 옵션 무시하고 '상품명'으로만 합산해서 TOP 3 추출
                    v_items = filtered_grouped[filtered_grouped['공급처'] == v_name]
                    v_top_merged = v_items.groupby('상품명')['최종잔량'].sum().sort_values(ascending=False).head(3)
                    
                    st.caption("🔥 TOP 3 상품 (통합)")
                    if not v_top_merged.empty:
                        for rank, (s_name, s_qty) in enumerate(v_top_merged.items()):
                            # 중복 없는 상품명과 통합 수량 표시
                            st.markdown(f"{rank+1}. {s_name} **({int(s_qty)}장)**")
                    else:
                        st.write("데이터 없음")
                    st.write("") 

    st.divider()

    # [4] 상세 표 (기존 유지)
    display_df = filtered_grouped.sort_values(by=['날짜_dt', '최종잔량'], ascending=[False, False])
    target_cols = ['날짜', '공급처', '상품명', '옵션', '최종잔량', '추가발주', '입고수량', '최종메모']
    
    st.dataframe(
        display_df[target_cols].rename(columns={'최종메모': '최근 처리내역(메모)'}), 
        use_container_width=True, 
        hide_index=True
    )

    import io
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        display_df[target_cols].to_excel(writer, index=False, sheet_name='리오더현황')
    st.download_button(
        label="📥 실시간 현황 엑셀 다운로드", 
        data=output.getvalue(), 
        file_name=f"리오더현황_{datetime.now(KST).strftime('%m%d_%H%M')}.xlsx", 
        use_container_width=True
    )


# ------------------------------------------------------------------
# 5️⃣단계: 전체 히스토리 기록 (실시간 데이터 로드 보정판)
# ------------------------------------------------------------------
if st.session_state.get('analyzed') or st.session_state.get('show_step6'):
    st.session_state.show_step6 = True
    st.divider()
    st.header("📜 5단계: 전체 히스토리 기록")

    # 🚨 [수정] 데이터가 세션에 있더라도 무조건 최신 데이터를 시트에서 다시 읽어옴
    try:
        with st.spinner("⏳ 히스토리 기록 불러오는 중..."):
            sh = get_sheet()
            ws_hist = sh.worksheet("히스토리")
            raw_data = ws_hist.get_all_values()
            
            if len(raw_data) > 1:
                cols_5 = [c.strip() for c in raw_data[0]]
                h_df = pd.DataFrame(raw_data[1:], columns=cols_5)
                
                # 컬럼 중복 제거 및 이름 통일 (비고/메모 통합)
                h_df = h_df.loc[:, ~h_df.columns.duplicated()]
                if '메모' in h_df.columns:
                    h_df.rename(columns={'메모': '비고(처리내역)'}, inplace=True)
                elif '비고' in h_df.columns:
                    h_df.rename(columns={'비고': '비고(처리내역)'}, inplace=True)
                
                # 최신 데이터를 세션에 강제 업데이트
                st.session_state.db_history = h_df
            else:
                st.session_state.db_history = pd.DataFrame()
    except Exception as e:
        st.error(f"히스토리 로드 실패: {e}")
        st.session_state.db_history = pd.DataFrame()

    # 실제 화면에 뿌릴 데이터 복사
    m_df_5 = st.session_state.get('db_history', pd.DataFrame()).copy()
    
    if not m_df_5.empty:
        # 날짜 컬럼 처리
        d_col = next((c for c in m_df_5.columns if '날짜' in c or '시간' in c), m_df_5.columns[0])
        m_df_5['날짜_dt'] = pd.to_datetime(m_df_5[d_col], errors='coerce', format='mixed')
        m_df_5['날짜_only'] = m_df_5['날짜_dt'].dt.date
        
        # [UI] 필터 레이아웃 (사장님 원본 유지)
        c1, c2, c3 = st.columns([1.5, 1.5, 1.2]) 
        with c1: 
            today_val = datetime.now(KST).date()
            # 날짜 범위 선택 (기본값 오늘)
            sel_dates_5 = st.date_input("📅 조회 날짜 범위", [today_val, today_val], key="h_date_vSplit_final")

        # 날짜 필터링 적용
        if isinstance(sel_dates_5, (list, tuple)) and len(sel_dates_5) == 2:
            period_df = m_df_5[(m_df_5['날짜_only'] >= sel_dates_5[0]) & (m_df_5['날짜_only'] <= sel_dates_5[1])]
        else:
            # 날짜 한 개만 선택된 경우 처리
            s_date = sel_dates_5[0] if isinstance(sel_dates_5, (list, tuple)) else sel_dates_5
            period_df = m_df_5[m_df_5['날짜_only'] == s_date]

        with c2: 
            h_name_5 = st.text_input("🔍 상품명/옵션 검색", key="h_name_vSplit_final")
        with c3:
            # 회차(시간) 선택 옵션 구성
            t_opts = ["전체 회차"] + sorted(period_df['날짜_dt'].dropna().dt.strftime('%Y-%m-%d %H:%M:%S').unique(), reverse=True)
            h_time_5 = st.selectbox("⏰ 저장 회차 선택", t_opts, key="h_time_vSplit_final")

        # 검색어 및 회차 필터 최종 적용
        df_dis = period_df.copy()
        if h_name_5:
            # 상품명이나 옵션 컬럼이 있는 경우만 검색 (사장님 장부 컬럼 기준)
            search_mask = df_dis.apply(lambda r: h_name_5.lower() in str(r).lower(), axis=1)
            df_dis = df_dis[search_mask]
            
        if h_time_5 != "전체 회차":
            df_dis = df_dis[df_dis['날짜_dt'].dt.strftime('%Y-%m-%d %H:%M:%S') == h_time_5]

        # 5단계 데이터 표 출력
        st.dataframe(
            df_dis.sort_values(by='날짜_dt', ascending=False).drop(columns=['날짜_dt', '날짜_only'], errors='ignore'), 
            use_container_width=True, 
            hide_index=True
        )
    else:
        st.info("현재 조회된 히스토리 기록이 없습니다. 먼저 저장을 진행해 주세요.")

    # 🚨 [핵심] 5단계 바로 아래 6단계를 호출하여 한 화면에 이어서 출력
    render_step6()
