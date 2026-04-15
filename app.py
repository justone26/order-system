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
# 1️⃣단계: 파일 업로드 및 데이터 로드 (파일까지 한 번에 초기화)
# ------------------------------------------------------------------
st.header("1️⃣ 파일 업로드 및 데이터 로드")

# 1. 파일 업로더용 리셋 키 설정 (최상단에 위치)
if 'uploader_key' not in st.session_state:
    st.session_state.uploader_key = 0

# 2. ✅ 업로더에 key를 부여해서 제어 가능하게 만듦
# 사장님, 여기서 key 뒤에 숫자가 바뀌면 파일이 싹 날아갑니다.
up_file = st.file_uploader(
    "엑셀 파일을 업로드하세요.", 
    type=['xlsx', 'xls'],
    key=f"file_uploader_{st.session_state.uploader_key}"
)

# 3. ✅ 초기화 버튼: 이제 파일 'X' 안 눌러도 됩니다!
if st.button("🔄 현재 화면 데이터 초기화", use_container_width=True):
    # 리셋 키를 제외한 모든 세션 상태 삭제
    for key in list(st.session_state.keys()):
        if key != 'uploader_key':
            del st.session_state[key]
    
    # 🚨 리셋 키 값을 올려서 파일 업로더를 강제로 비움
    st.session_state.uploader_key += 1
    
    st.success("✅ 파일과 모든 데이터가 초기화되었습니다.")
    time.sleep(0.5)
    st.rerun()

# 4. 파일 로드 로직 (기존 사장님 로직 유지)
if up_file is not None:
    if 'df_raw' not in st.session_state:
        try:
            temp_df = pd.read_excel(up_file)
            st.session_state.df_raw = temp_df
            st.session_state.analyzed = False
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
# 4️⃣단계: 입고 관리 및 최종 저장 (기능 100% 유지 + 동기화 보강)
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.header("📊 4단계: 입고 관리 및 최종 발주 확정")
    
    p = st.session_state.p
    
    if 'df_final' not in st.session_state:
        st.error("데이터가 없습니다. 이전 단계를 먼저 진행해 주세요.")
        st.stop()

    # [기능유지] 사장님표 상태 필터 및 검색 UI
    f1, f2 = st.columns([1, 2])
    with f1: 
        f_mode = st.selectbox("🚦 상태 필터", ["전체보기", "🚨 발주필요(세트)", "✅ 정상", "🚫 품절"], index=1)
    with f2: 
        s_query = st.text_input("🔍 검색 (상품명/옵션)")

    # [기능유지] 필터링 로직
    df_temp = st.session_state.df_final.copy()
    if f_mode == "🚨 발주필요(세트)":
        need_items = df_temp[(df_temp['상태'] != "🚫 품절") & (df_temp['권장발주수량'] > 0)][p['it']].unique()
        df_temp = df_temp[df_temp[p['it']].isin(need_items)]
    elif f_mode != "전체보기":
        df_temp = df_temp[df_temp['상태'] == f_mode]
        
    if s_query:
        df_temp = df_temp[df_temp[p['it']].str.contains(s_query, case=False) | 
                           df_temp[p['op']].str.contains(s_query, case=False)]

    # [기능유지] 컬럼 구성
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
                p['t3']: st.column_config.NumberColumn("3일판매", disabled=True),
                '일판매량': st.column_config.NumberColumn("일평균", disabled=True),
                '권장발주수량': st.column_config.NumberColumn("권장수량", disabled=True),
                '비고(처리내역)': st.column_config.TextColumn("📝 비고(처리내역)")
            }
        )
        btn_save = st.form_submit_button("🚀 최종 데이터 저장 및 시트 전송", use_container_width=True, type="primary")

    if btn_save:
        changed_rows = edited_df[(edited_df['입고차감'] > 0) | (edited_df['추가발주'] > 0)].copy()
        
        if not changed_rows.empty:
            with st.spinner("🚀 시트 전송 및 화면 데이터 동기화 중..."):
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
                        
                        m_parts = []
                        # ✨ 수정: 입고에도 날짜가 찍히도록 보강
                        if q_val > 0: m_parts.append(f"{time_short} {q_val}발주")
                        if i_val > 0: m_parts.append(f"{time_short} {i_val}입고")
                        
                        if m_parts:
                            auto_memo = f"[{' '.join(m_parts)}]"
                            final_memo = f"{auto_memo} {user_memo}".strip()
                        else:
                            final_memo = user_memo
                        
                        rows_qty.append([now_s, r[p['vn']], r[p['it']], r[p['op']], r[p['vi']], int(r['기존리오더']), q_val, i_val, final_memo])
                        rows_hist.append([now_s, r[p['vn']], r[p['it']], r[p['op']], r[p['vi']], r[p['av']], r['기존리오더'], i_val, q_val, r['권장발주수량'], final_memo])

                        # 화면 데이터 실시간 반영 (기능 유지)
                        mask = (st.session_state.df_final[p['it']] == r[p['it']]) & (st.session_state.df_final[p['op']] == r[p['op']])
                        st.session_state.df_final.loc[mask, '기존리오더'] = max(0, int(r['기존리오더']) + q_val - i_val)
                        st.session_state.df_final.loc[mask, '입고차감'] = 0
                        st.session_state.df_final.loc[mask, '추가발주'] = 0

                    if rows_qty: ws_qty.append_rows(rows_qty, value_input_option='USER_ENTERED')
                    if rows_hist: ws_hist.append_rows(rows_hist, value_input_option='USER_ENTERED')
                    
                    # 🚨 [중요] 5, 6단계 강제 새로고침 로직
                    st.session_state.show_step6 = True
                    if 'db_history' in st.session_state: del st.session_state.db_history
                    if 'master_log' in st.session_state: del st.session_state.master_log
                    st.cache_data.clear()

                    st.success(f"✅ 저장 성공! 히스토리와 현황판이 업데이트되었습니다.")
                    time.sleep(1)
                    st.rerun() 
                    
                except Exception as e:
                    st.error(f"저장 중 오류 발생: {e}")
        else:
            st.warning("⚠️ 저장할 변경 내역이 없습니다.")

# ------------------------------------------------------------------
# 6️⃣단계: 리오더 현황판 (수량 옆으로 밀착 + 공간 최적화 최종)
# ------------------------------------------------------------------
def render_step6():
    if not (st.session_state.get('analyzed') or st.session_state.get('show_step6')):
        return

    st.markdown("---")
    st.markdown("### 📈 6단계: 실시간 리오더 현황판 (상품별 통합)")
    
    try:
        sh = get_sheet()
        ws_qty = sh.worksheet("발주기록")
        df_log = pd.DataFrame(ws_qty.get_all_records())
        if df_log.empty: return
        
        # 숫자 및 날짜 데이터 전처리
        for col in ['기존리오더', '추가발주', '입고수량']:
            df_log[col] = pd.to_numeric(df_log[col], errors='coerce').fillna(0)
        df_log['날짜_dt'] = pd.to_datetime(df_log['날짜'], errors='coerce', format='mixed')
    except Exception as e:
        st.error(f"데이터 로드 중 오류: {e}"); return

    # [1] UI 레이아웃 - 상단 정렬
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

    # [2] 스마트 FIFO 미입고 계산 함수
    def get_fifo_pending(group):
        total_in = group['입고수량'].sum()
        pending_list = []
        sorted_orders = group[group['추가발주'] > 0].sort_values('날짜_dt', ascending=True)
        remaining_in = total_in
        for _, row in sorted_orders.iterrows():
            order_amt = row['추가발주']
            if remaining_in >= order_amt:
                remaining_in -= order_amt
            else:
                actual_pending = order_amt - remaining_in
                pending_list.append(f"{row['날짜_dt'].strftime('%m/%d')} {int(actual_pending)}장")
                remaining_in = 0
        return " / ".join(pending_list)

    # [3] 데이터 그룹화
    grouped = df_log.groupby(['공급처', '상품명', '옵션']).apply(lambda x: pd.Series({
        '날짜': x['날짜_dt'].max(),
        '기존리오더': x['기존리오더'].sum(),
        '추가발주': x['추가발주'].sum(),
        '입고수량': x['입고수량'].sum(),
        '미입고히스토리': get_fifo_pending(x),
        '수기메모': " / ".join(set([str(i) for i in x['메모'] if str(i).strip() != ""]))
    })).reset_index()

    grouped['최종잔량'] = grouped['기존리오더'] + grouped['추가발주'] - grouped['입고수량']
    grouped['최종메모'] = grouped.apply(
        lambda x: f"[{x['미입고히스토리']}] {x['수기메모']}".strip() if x['미입고히스토리'] else x['수기메모'], axis=1
    )

    filtered_grouped = grouped.copy()
    if sel_s: filtered_grouped = filtered_grouped[filtered_grouped['상품명'].str.contains(sel_s, case=False)]
    if sel_v != "전체 공급처": filtered_grouped = filtered_grouped[filtered_grouped['공급처'] == sel_v]

    # [4] 🏢 업체별 요약 (수량 옆으로 배치)
    st.markdown("#### 🏢 업체별 미입고 및 주요 상품")
    v_sum = filtered_grouped.groupby('공급처')['최종잔량'].sum().reset_index()
    v_sum = v_sum[v_sum['최종잔량'] > 0].sort_values('최종잔량', ascending=False)
    
    if not v_sum.empty:
        v_cols = st.columns(min(len(v_sum), 4))
        for i, (idx, row) in enumerate(v_sum.iterrows()):
            if i < 4:
                v_name = row['공급처']
                with v_cols[i]:
                    st.metric(v_name, f"{int(row['최종잔량'])}개 잔량")
                    
                    v_items = filtered_grouped[filtered_grouped['공급처'] == v_name]
                    v_top_combined = v_items.groupby('상품명')['최종잔량'].sum().sort_values(ascending=False).head(3)
                    
                    st.caption("🔥 TOP 3 상품 (통합)")
                    for rank, (s_name, s_qty) in enumerate(v_top_combined.items()):
                        # 🚨 [수정 포인트] 상품명 바로 옆에 수량을 붙여서 공간 절약
                        st.markdown(f"{rank+1}. {s_name} **({int(s_qty)}장)**")
                    st.write("") 

    st.divider()

    # [5] 상세 표 및 엑셀 다운로드
    display_df = filtered_grouped.sort_values(by=['날짜', '최종잔량'], ascending=[False, False])
    target_cols = ['날짜', '공급처', '상품명', '옵션', '최종잔량', '추가발주', '입고수량', '최종메모']
    st.dataframe(display_df[target_cols].rename(columns={'최종메모': '미입고 상세기록'}), use_container_width=True, hide_index=True)

    import io
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        display_df[target_cols].to_excel(writer, index=False, sheet_name='리오더현황')
    st.download_button(label="📥 실시간 현황 엑셀 다운로드", data=output.getvalue(), file_name=f"리오더현황_{datetime.now(KST).strftime('%m%d_%H%M')}.xlsx", use_container_width=True)


# ------------------------------------------------------------------
# 5️⃣단계: 전체 히스토리 기록 (실행문)
# ------------------------------------------------------------------
if st.session_state.get('analyzed') or st.session_state.get('show_step6'):
    st.session_state.show_step6 = True
    st.divider()
    st.header("📜 5단계: 전체 히스토리 기록")

    # 데이터 로드 (중복 제거 및 메모 컬럼 통일)
    if 'db_history' not in st.session_state:
        try:
            sh = get_sheet()
            ws_hist = sh.worksheet("히스토리")
            raw_data = ws_hist.get_all_values()
            if len(raw_data) > 1:
                cols_5 = [c.strip() for c in raw_data[0]]
                h_df = pd.DataFrame(raw_data[1:], columns=cols_5)
                h_df = h_df.loc[:, ~h_df.columns.duplicated()]
                h_df.rename(columns={'메모': '비고(처리내역)', '비고': '비고(처리내역)'}, errors='ignore', inplace=True)
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

        # 당일 데이터 추출 및 회차 필터 생성
        if isinstance(sel_dates_5, (list, tuple)) and len(sel_dates_5) == 2:
            period_df = m_df_5[(m_df_5['날짜_only'] >= sel_dates_5[0]) & (m_df_5['날짜_only'] <= sel_dates_5[1])]
        else:
            period_df = m_df_5[m_df_5['날짜_only'] == sel_dates_5]

        with c2: h_name_5 = st.text_input("🔍 상품명/옵션 검색", key="h_name_vSplit_01")
        with c3:
            t_opts = ["전체 회차"] + sorted(period_df['날짜_dt'].dropna().dt.strftime('%Y-%m-%d %H:%M:%S').unique(), reverse=True)
            h_time_5 = st.selectbox("⏰ 저장 회차 선택", t_opts, key="h_time_vSplit_01")

        df_dis = period_df.copy()
        if h_name_5: df_dis = df_dis[df_dis.apply(lambda r: h_name_5.lower() in str(r).lower(), axis=1)]
        if h_time_5 != "전체 회차": df_dis = df_dis[df_dis['날짜_dt'].dt.strftime('%Y-%m-%d %H:%M:%S') == h_time_5]

        st.dataframe(df_dis.sort_values(by='날짜_dt', ascending=False).drop(columns=['날짜_dt', '날짜_only'], errors='ignore'), use_container_width=True, hide_index=True)

    # 🚨 [중요] 5단계가 끝나는 지점에서 6단계를 호출합니다.
    # 이 줄이 실행될 때 위에 정의된 render_step6 함수를 찾아갑니다.
    render_step6()
