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
# 1️⃣단계: 파일 업로드 및 데이터 로드 (마지막 잔액 추출 방식)
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
            
            with st.spinner("🔄 구글 시트에서 마지막 잔액(통장 잔액) 동기화 중..."):
                sh = get_sheet()
                ws = sh.worksheet("발주기록")
                raw_records = ws.get_all_values()
                
                reorder_balance_map = {} # 🚨 잔액 저장소
                
                if len(raw_records) > 1:
                    df_rec = pd.DataFrame(raw_records[1:], columns=[h.strip() for h in raw_records[0]])
                    def c_func(t): return "".join(str(t).split()).upper()
                    df_rec['key'] = df_rec['상품명'].apply(c_func) + df_rec['옵션'].apply(c_func)
                    
                    # 🚨 [사장님 로직] 전체 합산이 아니라, 각 상품의 '가장 마지막 줄' 기존리오더를 가져옴
                    # drop_duplicates를 사용하여 가장 최근(마지막) 데이터만 남김
                    last_entries = df_rec.drop_duplicates('key', keep='last')
                    for _, row in last_entries.iterrows():
                        val = str(row['기존리오더']).replace(',', '')
                        reorder_balance_map[row['key']] = int(pd.to_numeric(val, errors='coerce') or 0)
                
                # 3단계에서 쓸 정답지 세션 저장
                st.session_state.reorder_ans = reorder_balance_map
                    
            st.success("✅ 엑셀 및 장부 잔액 동기화 완료!")
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
# 3️⃣단계: 분석 설정 및 실행 (사전 로드된 잔액 적용)
# ------------------------------------------------------------------
if 'df_raw' in st.session_state:
    st.divider()
    st.subheader("⚙️ 3️⃣단계: 분석 설정 및 실행")

    clt, css = st.columns(2)
    with clt: lt = st.number_input("리드타임 (일)", value=7, key="input_lt")
    with css: ss = st.number_input("안전재고 (일 수)", value=3, key="input_ss")

    if st.button("🚀 분석 실행", type="primary", use_container_width=True):
        try:
            p_map = {
                'so': st.session_state.get('sel_so'), 'it': st.session_state.get('sel_it'),
                'op': st.session_state.get('sel_op'), 'vn': st.session_state.get('sel_vn'),
                'vi': st.session_state.get('sel_vi'), 'av': st.session_state.get('sel_av'),
                't3': st.session_state.get('sel_t3'), 't7': st.session_state.get('sel_t7'),
                'rd': st.session_state.get('sel_rd'), 'lt': lt, 'ss': ss
            }
            st.session_state.p = p_map

            with st.spinner("📊 통장 잔액 기준으로 권장수량 재계산 중..."):
                df = st.session_state.df_raw.copy()
                def c_func(t): return "".join(str(t).split()).upper()
                df['key'] = df[p_map['it']].apply(c_func) + df[p_map['op']].apply(c_func)

                # 🚨 [정교화] 미리 로드한 상품별 마지막 잔액을 주입
                ans_map = st.session_state.get('reorder_ans', {})
                df['기존리오더'] = df['key'].map(ans_map).fillna(0).astype(int)

                # 수량 변환 및 일판매량 계산
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

                # 🚨 권장발주수량 공식 적용
                df['권장발주수량'] = ((df['일판매량'] * (lt + ss)) - (df[c_av] + df['기존리오더'])).clip(lower=0).astype(int)
                
                df['상태'] = df.apply(lambda r: "🚫 품절" if "품절" in str(r[p_map['so']]) else ("🚨 발주필요" if r['권장발주수량'] > 0 else "✅ 정상"), axis=1)
                df['입고차감'] = 0 ; df['추가발주'] = 0 ; df['비고(처리내역)'] = ""

                if 'key' in df.columns: df = df.drop(columns=['key'])
                st.session_state.df_final = df
                st.session_state.analyzed = True
                st.cache_data.clear()
                st.success(f"✅ 분석 완료 (통장 잔액 반영됨)")
                st.rerun()

        except Exception as e:
            st.error(f"⚠️ 분석 오류: {e}")


# ------------------------------------------------------------------
# 4️⃣단계: 입고 관리 및 최종 저장 (모든 누락 컬럼 완전 복구 버전)
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.header("📊 4단계: 입고 관리 및 최종 발주 확정")
    
    p = st.session_state.p
    df_all = st.session_state.df_final.copy()

    # [1] 필터 UI
    f1, f2 = st.columns([1, 2])
    with f1: 
        f_mode = st.selectbox(
            "🚦 상태 필터", 
            ["🚨 발주필요(세트)", "✅ 정상(품절제외)", "🚫 품절건만 보기", "전체보기"], 
            index=0
        )
    with f2: 
        s_query = st.text_input("🔍 검색 (상품명/옵션/공급처상품명)", key="s4_final_fix_search")

    # [2] 필터링 로직
    df_temp = df_all.copy()

    if f_mode == "🚨 발주필요(세트)":
        df_temp = df_temp[~df_temp['상태'].str.contains("품절", na=False)]
        need_items = df_temp[df_temp['권장발주수량'] > 0][p['it']].unique()
        df_temp = df_temp[df_temp[p['it']].isin(need_items)]
        
    elif f_mode == "✅ 정상(품절제외)":
        df_temp = df_temp[(df_temp['상태'] == "✅ 정상") & (~df_temp['상태'].str.contains("품절", na=False))]
        
    elif f_mode == "🚫 품절건만 보기":
        df_temp = df_temp[df_temp['상태'].str.contains("품절", na=False)]
        
    if s_query:
        df_temp = df_temp[
            df_temp[p['it']].str.contains(s_query, case=False, na=False) | 
            df_temp[p['op']].str.contains(s_query, case=False, na=False) |
            df_temp[p['vi']].str.contains(s_query, case=False, na=False)
        ]

    # [3] 사장님 UI 컬럼 순서 (3일판매량 p['t3'] 추가 복구)
    # 순서: 상태 | 공급처 | 상품명 | 옵션 | 공급처상품명 | 가용재고 | 기존리오더 | 입고차감 | 추가발주 | 3일판매량 | 7일판매량(일판매) | 권장수량 | 비고
    disp_cols = [
        '상태', p['vn'], p['it'], p['op'], p['vi'], p['av'], 
        '기존리오더', '입고차감', '추가발주', p['t3'], '일판매량', '권장발주수량', '비고(처리내역)'
    ]
    
    # 실제 데이터프레임에 있는 컬럼만 필터링
    disp_cols = [c for c in disp_cols if c in df_temp.columns]
    
    with st.form("final_form"):
        edited_df = st.data_editor(
            df_temp[disp_cols], 
            use_container_width=True, 
            hide_index=True,
            column_config={
                '상태': st.column_config.TextColumn("상태", disabled=True),
                p['vn']: st.column_config.TextColumn("공급처", disabled=True),
                p['it']: st.column_config.TextColumn("상품명", disabled=True),
                p['op']: st.column_config.TextColumn("옵션", disabled=True),
                p['vi']: st.column_config.TextColumn("공급처상품명", disabled=True),
                p['av']: st.column_config.NumberColumn("가용재고", disabled=True, format="%d"),
                '기존리오더': st.column_config.NumberColumn("기존리오더", disabled=True, format="%d"),
                '입고차감': st.column_config.NumberColumn("📥 입고(-)", min_value=0), 
                '추가발주': st.column_config.NumberColumn("➕ 발주(+)", min_value=0),
                p['t3']: st.column_config.NumberColumn("3일판매", disabled=True), # 👈 복구 완료
                '일판매량': st.column_config.NumberColumn("평균일판매", disabled=True),
                '권장발주수량': st.column_config.NumberColumn("권장수량", disabled=True),
            }
        )
        btn_save = st.form_submit_button("🚀 최종 데이터 저장 및 기존리오더 업데이트", use_container_width=True, type="primary")

    if btn_save:
        changed_rows = edited_df[(edited_df['입고차감'] > 0) | (edited_df['추가발주'] > 0)].copy()
        
        if not changed_rows.empty:
            with st.spinner("🚀 기존리오더 계산 중..."):
                try:
                    sh = get_sheet()
                    ws_qty = sh.worksheet("발주기록")
                    now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
                    
                    rows_to_save = []
                    for _, r in changed_rows.iterrows():
                        # 계산 로직: 기존리오더 + 추가발주 - 입고차감
                        new_reorder_val = int(r['기존리오더']) + int(r['추가발주']) - int(r['입고차감'])
                        
                        rows_to_save.append([
                            now_s, r[p['vn']], r[p['it']], r[p['op']], 
                            new_reorder_val, int(r['추가발주']), int(r['입고차감']), r['비고(처리내역)']
                        ])

                    if rows_to_save:
                        ws_qty.append_rows(rows_to_save, value_input_option='USER_ENTERED')
                        st.success(f"✅ {len(rows_to_save)}건 저장 완료! 기존리오더가 업데이트되었습니다.")
                        st.cache_data.clear()
                        time.sleep(1)
                        st.rerun()
                except Exception as e:
                    st.error(f"⚠️ 저장 오류: {e}")
        else:
            st.warning("⚠️ 입력된 수량이 없습니다.")


# ------------------------------------------------------------------
# 6️⃣단계: 리오더 현황판 (최신 잔액(통장방식) 반영 버전)
# ------------------------------------------------------------------
def render_step6():
    if not (st.session_state.get('analyzed') or st.session_state.get('show_step6')):
        return

    st.markdown("---")
    st.markdown("### 📈 6단계: 실시간 리오더 현황판 (통장 잔액 기준)")
    
    try:
        sh = get_sheet()
        ws_qty = sh.worksheet("발주기록")
        # get_all_records() 대신 get_all_values()로 안전하게 로드
        data = ws_qty.get_all_values()
        if len(data) <= 1: return
        
        df_log = pd.DataFrame(data[1:], columns=[h.strip() for h in data[0]])
        
        # 숫자 데이터 전처리 (에러 방지)
        for col in ['기존리오더', '추가발주', '입고수량']:
            df_log[col] = pd.to_numeric(df_log[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        
        df_log['날짜_dt'] = pd.to_datetime(df_log['날짜'], errors='coerce', format='mixed')
    except Exception as e:
        st.error(f"데이터 로드 중 오류: {e}"); return

    # [1] UI 레이아웃
    c1, c2, c3 = st.columns([1, 2, 1.5])
    with c1:
        st.markdown("<p style='margin-bottom: 8px; font-size: 14px;'>🔄 데이터 갱신</p>", unsafe_allow_html=True)
        if st.button("최신 자료 업데이트", use_container_width=True, key="btn_update_compact"):
            st.cache_data.clear()
            st.rerun()
    with c2:
        sel_s = st.text_input("🔍 상품명 검색", placeholder="상품명 입력", key="s6_search_compact")
    with c3:
        v_list = ["전체 공급처"] + sorted(df_log['공급처'].unique().tolist())
        sel_v = st.selectbox("🏭 공급처 필터", v_list, key="s6_vendor_compact")

    # [2] 🚨 사장님 로직 적용: 상품별 '마지막 행'의 잔액이 곧 최종 수량
    def c(t): return "".join(str(t).split()).upper()
    df_log['key'] = df_log['상품명'].apply(c) + df_log['옵션'].apply(c)
    
    # 중복 제거 시 'keep=last'를 사용하여 가장 최신 잔액행만 남김
    # 추가발주와 입고수량은 '전체 기간 합계'가 아니라 '최근 내역'임을 참고
    grouped = df_log.sort_values('날짜_dt').drop_duplicates('key', keep='last').copy()
    
    # 최종잔량 = 장부의 '기존리오더' 컬럼 그 자체 (이미 4단계에서 계산해서 저장했으므로)
    grouped['최종잔량'] = grouped['기존리오더']
    
    # 메모 정리 (최근 5개 메모만 합쳐서 노출)
    memo_map = df_log.groupby('key')['메모'].apply(lambda x: " / ".join([str(i) for i in x.tail(5) if str(i).strip() != ""])).to_dict()
    grouped['최종메모'] = grouped['key'].map(memo_map)

    filtered_grouped = grouped.copy()
    if sel_s: filtered_grouped = filtered_grouped[filtered_grouped['상품명'].str.contains(sel_s, case=False)]
    if sel_v != "전체 공급처": filtered_grouped = filtered_grouped[filtered_grouped['공급처'] == sel_v]

    # [3] 업체별 요약
    st.markdown("#### 🏢 업체별 미입고 잔량 요약")
    v_sum = filtered_grouped.groupby('공급처')['최종잔량'].sum().reset_index()
    v_sum = v_sum[v_sum['최종잔량'] > 0].sort_values('최종잔량', ascending=False)
    
    if not v_sum.empty:
        v_cols = st.columns(min(len(v_sum), 4))
        for i, (idx, row) in enumerate(v_sum.iterrows()):
            if i < 4:
                v_name = row['공급처']
                with v_cols[i]:
                    st.metric(v_name, f"{int(row['최종잔량'])}개")
                    v_top = filtered_grouped[filtered_grouped['공급처'] == v_name].sort_values('최종잔량', ascending=False).head(3)
                    for rank, (_, r) in enumerate(v_top.iterrows()):
                        st.markdown(f"{rank+1}. {r['상품명']} **({int(r['최종잔량'])}장)**")

    st.divider()

    # [4] 상세 표
    display_df = filtered_grouped.sort_values(by=['날짜_dt', '최종잔량'], ascending=[False, False])
    target_cols = ['날짜', '공급처', '상품명', '옵션', '최종잔량', '추가발주', '입고수량', '최종메모']
    st.dataframe(display_df[target_cols].rename(columns={'최종메모': '최근 처리내역(메모)'}), use_container_width=True, hide_index=True)


# ------------------------------------------------------------------
# 5️⃣단계: 전체 히스토리 기록 (로그 조회용)
# ------------------------------------------------------------------
if st.session_state.get('analyzed') or st.session_state.get('show_step6'):
    st.session_state.show_step6 = True
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
                # 컬럼 중복 제거 및 이름 통일
                h_df = h_df.loc[:, ~h_df.columns.duplicated()]
                st.session_state.db_history = h_df
            else: st.session_state.db_history = pd.DataFrame()
        except: st.session_state.db_history = pd.DataFrame()

    m_df_5 = st.session_state.get('db_history', pd.DataFrame()).copy()
    
    if not m_df_5.empty:
        d_col = next((c for c in m_df_5.columns if '날짜' in c or '시간' in c), m_df_5.columns[0])
        m_df_5['날짜_dt'] = pd.to_datetime(m_df_5[d_col], errors='coerce', format='mixed')
        m_df_5['날짜_only'] = m_df_5['날짜_dt'].dt.date
        
        c1, c2, c3 = st.columns([1.5, 1.5, 1.2]) 
        with c1: 
            today_val = datetime.now(KST).date()
            sel_dates_5 = st.date_input("📅 조회 날짜 범위", [today_val, today_val], key="h_date_vSplit_01")
        with c2: h_name_5 = st.text_input("🔍 상품명/옵션 검색", key="h_name_vSplit_01")
        with c3:
            t_opts = ["전체 회차"] + sorted(m_df_5['날짜_dt'].dropna().dt.strftime('%Y-%m-%d %H:%M:%S').unique(), reverse=True)
            h_time_5 = st.selectbox("⏰ 저장 회차 선택", t_opts, key="h_time_vSplit_01")

        # 필터링
        df_dis = m_df_5.copy()
        if isinstance(sel_dates_5, (list, tuple)) and len(sel_dates_5) == 2:
            df_dis = df_dis[(df_dis['날짜_only'] >= sel_dates_5[0]) & (df_dis['날짜_only'] <= sel_dates_5[1])]
        if h_name_5: 
            df_dis = df_dis[df_dis['상품명'].str.contains(h_name_5, case=False, na=False) | 
                            df_dis['옵션'].str.contains(h_name_5, case=False, na=False)]
        if h_time_5 != "전체 회차": 
            df_dis = df_dis[df_dis['날짜_dt'].dt.strftime('%Y-%m-%d %H:%M:%S') == h_time_5]

        st.dataframe(df_dis.sort_values(by='날짜_dt', ascending=False).drop(columns=['날짜_dt', '날짜_only'], errors='ignore'), use_container_width=True, hide_index=True)

    # 🚨 수정된 6단계 호출 (최신 잔액 방식으로 렌더링됨)
    render_step6()
