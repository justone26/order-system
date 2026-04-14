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
    # 3️⃣단계: 분석 설정 및 실행 (발주/입고 실시간 잔량 계산 통합 버전)
    # ------------------------------------------------------------------
    st.divider()
    st.subheader("⚙️ 3️⃣단계: 분석 설정 및 실행")
    
    clt, css = st.columns(2)
    with clt:
        lt = st.number_input("리드타임 (일)", value=10, help="주문 후 입고까지 걸리는 평균 일수")
    with css:
        ss = st.number_input("안전재고 (일 수)", value=7, help="품절 방지를 위해 추가로 보유할 재고 일수")

    if st.button("🚀 분석 실행 / 실시간 장부 업데이트", type="primary", use_container_width=True):
        # 파라미터 저장
        st.session_state.p = {
            'so': sold_out, 'it': item, 'op': option, 'vn': vendor, 'vi': v_item_col,
            'av': avail, 't3': t3d, 't7': t1w, 'lt': lt, 'ss': ss, 'rd': reg_date
        }

        with st.spinner("📊 구글 시트 동기화 및 실시간 미입고 잔량 분석 중..."):
            try:
                # 1. 기초 데이터 복사
                df = st.session_state.df_raw.copy()
                today = datetime.now(KST).date()

                # 2. 구글 시트 데이터 로드 (get_all_values로 헤더 에러 방지)
                sh = get_sheet()
                
                # [발주기록 시트 읽기]
                ws_qty = sh.worksheet("발주기록")
                qty_data = ws_qty.get_all_values()
                if len(qty_data) > 1:
                    df_qty = pd.DataFrame(qty_data[1:], columns=[c.strip() for c in qty_data[0]])
                    df_qty = df_qty.loc[:, ~df_qty.columns.duplicated()] # 중복 컬럼 제거
                else:
                    df_qty = pd.DataFrame()

                # [입고기록 시트 읽기]
                ws_in = sh.worksheet("입고기록")
                in_data = ws_in.get_all_values()
                if len(in_data) > 1:
                    df_in = pd.DataFrame(in_data[1:], columns=[c.strip() for c in in_data[0]])
                    df_in = df_in.loc[:, ~df_in.columns.duplicated()] # 중복 컬럼 제거
                else:
                    df_in = pd.DataFrame()

                # 5, 6단계에서 활용하기 위해 세션에 발주기록 저장
                st.session_state.master_log = df_qty 

                # 3. 실시간 리오더 잔량(미입고) 계산
                r_map = {}
                if not df_qty.empty:
                    # 수량 컬럼 찾기 (이름이 달라도 유연하게 대응)
                    q_col = next((c for c in ['추가발주수량', '발주수량', '추가발주', '수량'] if c in df_qty.columns), None)
                    i_col = next((c for c in ['입고수량', '입고차감', '입고', '수량'] if c in df_in.columns), None)
                    
                    if q_col:
                        # 발주 합계 계산
                        df_qty[q_col] = pd.to_numeric(df_qty[q_col], errors='coerce').fillna(0)
                        qty_sum = df_qty.groupby(['상품명', '옵션'])[q_col].sum()

                        # 입고 합계 계산
                        if not df_in.empty and i_col:
                            df_in[i_col] = pd.to_numeric(df_in[i_col], errors='coerce').fillna(0)
                            in_sum = df_in.groupby(['상품명', '옵션'])[i_col].sum()
                        else:
                            in_sum = pd.Series()

                        # 최종 잔량 = 발주합계 - 입고합계
                        r_map = qty_sum.sub(in_sum, fill_value=0).clip(lower=0).to_dict()

                # 4. 분석 계산 로직 적용
                # 숫자 변환
                df[avail] = pd.to_numeric(df[avail], errors='coerce').fillna(0).astype(int)
                df[t1w] = pd.to_numeric(df[t1w], errors='coerce').fillna(0).astype(int)
                
                # 실시간 잔량을 '기존리오더' 컬럼에 매칭
                def get_reorder_val(row):
                    k = (str(row[item]).strip(), str(row[option]).strip())
                    return int(r_map.get(k, 0))
                
                df['기존리오더'] = df.apply(get_reorder_val, axis=1)

                # 일판매량 계산 (등록일 기준 가중치)
                def get_daily_avg(row):
                    try:
                        r_dt = pd.to_datetime(row[reg_date]).date()
                        days = max(1, min((today - r_dt).days, 7))
                        return int(round(pd.to_numeric(row[t1w]) / days, 0))
                    except: 
                        return int(round(pd.to_numeric(row[t1w]) / 7, 0))

                df['일판매량'] = df.apply(get_daily_avg, axis=1)
                
                # 권장발주수량 공식: (판매량 * 기간) - (창고재고 + 미입고잔량)
                df['권장발주수량'] = ((df['일판매량'] * (lt + ss)) - (df[avail] + df['기존리오더'])).clip(lower=0).astype(int)
                
                # 상태 판별
                df['상태'] = df.apply(lambda r: "🚫 품절" if "품절" in str(r[sold_out]) else ("🚨 발주필요" if r['권장발주수량'] > 0 else "✅ 정상"), axis=1)
                
                # 4단계용 추가 입력 컬럼 초기화
                df['입고차감'] = 0
                df['추가발주'] = 0
                df['비고(메모)'] = ""
                
                # 결과 세션 저장 및 완료 알림
                st.session_state.df_final = df
                st.session_state.analyzed = True
                st.success("✅ 분석 완료! 시트의 미입고 잔량이 실시간 반영되었습니다.")
                st.rerun()
                
            except Exception as e:
                st.error(f"분석 중 오류 발생: {e}")
                st.info("💡 팁: '발주기록'과 '입고기록' 시트의 첫 줄(헤더)에 빈 칸이나 중복된 이름이 없는지 확인해주세요.")


                
# ------------------------------------------------------------------
# 4️⃣단계: 입고 관리 및 최종 저장 (3중 시트 전송 버전)
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.header("📊 4단계: 입고 관리 및 최종 발주 확정")
    
    p = st.session_state.p
    df_disp = st.session_state.df_final.copy()
    
    f1, f2 = st.columns([1, 2])
    with f1: f_mode = st.selectbox("🚦 상태 필터", ["전체보기", "🚨 발주필요(세트)", "✅ 정상", "🚫 품절"], index=1)
    with f2: s_query = st.text_input("🔍 검색 (상품명/옵션)")

    # 필터 적용 로직
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
        btn_save = st.form_submit_button("🚀 최종 데이터 저장 및 시트 전송", use_container_width=True, type="primary")

    if btn_save:
        change_list = edited_df[(edited_df['입고차감'] > 0) | (edited_df['추가발주'] > 0)]
        if not change_list.empty:
            try:
                sh = get_sheet()
                # 시트 연결
                ws_in = sh.worksheet("입고기록")
                ws_qty = sh.worksheet("발주기록")
                ws_hist = sh.worksheet("히스토리")
                
                now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
                time_short = datetime.now(KST).strftime('%m/%d')
                
                rows_in, rows_qty, rows_hist = [], [], []

                for _, r in change_list.iterrows():
                    # 1. 메모 자동 생성 (6단계 방식)
                    q_val = int(r['추가발주'])
                    i_val = int(r['입고차감'])
                    user_memo = str(r['비고(메모)']).strip() if r['비고(메모)'] and str(r['비고(메모)']) != "None" else ""
                    
                    parts = []
                    if q_val > 0: parts.append(f"{q_val}발주")
                    if i_val > 0: parts.append(f"-{i_val}입고")
                    auto_memo = f"[{time_short} " + " ".join(parts) + "]"
                    final_memo = f"{auto_memo} {user_memo}".strip()
                    
                    # 2. 입고기록 시트용
                    if i_val > 0:
                        rows_in.append([now_s, r[p['vn']], r[p['it']], r[p['op']], r[p['vi']], i_val, final_memo])
                    
                    # 3. 발주기록 시트용
                    if q_val > 0:
                        rows_qty.append([now_s, r[p['vn']], r[p['it']], r[p['op']], r[p['vi']], q_val, final_memo])
                    
                    # 4. 전체 히스토리 시트용 (무조건 누적)
                    rows_hist.append([
                        now_s, r[p['vn']], r[p['it']], r[p['op']], r[p['vi']], 
                        r[p['av']], r['기존리오더'], i_val, q_val, r['권장발주수량'], final_memo
                    ])
                
                # 시트 전송
                if rows_in: ws_in.append_rows(rows_in)
                if rows_qty: ws_qty.append_rows(rows_qty)
                if rows_hist: ws_hist.append_rows(rows_hist)
                
                st.success(f"✅ 저장 완료! (입고 {len(rows_in)}건 / 발주 {len(rows_qty)}건 / 히스토리 {len(rows_hist)}건)")
                
                # 저장 후 동기화를 위해 세션 삭제
                if 'master_log' in st.session_state: del st.session_state.master_log
                time.sleep(1)
                st.rerun()
            except Exception as e: st.error(f"저장 실패: {e} (시트 이름을 확인하세요)")
                

# ------------------------------------------------------------------
# 5️⃣단계: 전체 히스토리 기록 (시트 실시간 로드 버전)
# ------------------------------------------------------------------
st.divider()
st.header("📜 5단계: 전체 히스토리 기록")

# 5단계는 항상 시트의 최신 데이터를 가져오도록 설계
if st.button("🔄 히스토리 불러오기/새로고침", use_container_width=True):
    try:
        sh = get_sheet()
        ws_hist = sh.worksheet("히스토리")
        data = ws_hist.get_all_records()
        if data:
            h_df = pd.DataFrame(data)
            h_df.columns = [c.strip() for c in h_df.columns]
            st.session_state.db_history = h_df
            st.success("✅ 시트에서 히스토리를 성공적으로 가져왔습니다.")
        else:
            st.warning("기록된 히스토리가 없습니다.")
    except Exception as e:
        st.error(f"히스토리 로드 실패: {e}")

if 'db_history' in st.session_state:
    m_df = st.session_state.db_history.copy()
    
    # 날짜 필터 준비
    m_df['날짜'] = pd.to_datetime(m_df['날짜'], errors='coerce')
    m_df['날짜_only'] = m_df['날짜'].dt.date
    valid_df = m_df.dropna(subset=['날짜'])
    
    if not valid_df.empty:
        c1, c2, c3 = st.columns(3)
        with c1: 
            sel_dates = st.date_input("📅 조회 날짜", [valid_df['날짜_only'].min(), valid_df['날짜_only'].max()], key="h_date_v10")
        with c2: 
            h_name = st.text_input("🔍 상품명 검색", key="h_name_v10")
        with c3: 
            t_options = ["전체 회차"] + [t.strftime('%Y-%m-%d %H:%M:%S') for t in sorted(valid_df['날짜'].unique(), reverse=True)]
            h_time = st.selectbox("⏰ 회차(시간) 선택", t_options, key="h_time_v10")

        # 필터 적용
        df_5 = valid_df.copy()
        if isinstance(sel_dates, (list, tuple)) and len(sel_dates) == 2:
            df_5 = df_5[(df_5['날짜_only'] >= sel_dates[0]) & (df_5['날짜_only'] <= sel_dates[1])]
        if h_name:
            df_5 = df_5[df_5['상품명'].astype(str).str.contains(h_name, case=False)]
        if h_time != "전체 회차":
            df_5 = df_5[df_5['날짜'].dt.strftime('%Y-%m-%d %H:%M:%S') == h_time]

        # 컬럼 순서 및 보정 (사용자 요청 기반)
        target_cols = ['날짜', '업체명', '상품명', '옵션', '공급처상품명', '가용재고', '기존리오더', '입고수량', '추가발주수량', '권장발주', '메모']
        for col in target_cols:
            if col not in df_5.columns: df_5[col] = 0

        st.dataframe(df_5[target_cols].sort_values('날짜', ascending=False), use_container_width=True, hide_index=True)
    else:
        st.info("조회할 수 있는 유효한 날짜 데이터가 없습니다.")
else:
    st.info("💡 위 '히스토리 불러오기' 버튼을 눌러 과거 기록을 확인하세요.")

# ------------------------------------------------------------------
# 6️⃣단계: 실시간 리오더 현황판 (입고/발주 시트 기반 실시간 계산)
# ------------------------------------------------------------------
if st.session_state.get('analyzed'):
    st.divider()
    st.header("📊 6단계: 실시간 리오더 현황판 (미입고 잔량)")

    if st.button("🔄 실시간 시트 데이터 동기화", use_container_width=True, key="sync_v11"):
        with st.spinner("📡 입고/발주 장부를 분석 중..."):
            try:
                sh = get_sheet()
                # 두 시트 데이터를 모두 가져옴
                df_in = pd.DataFrame(sh.worksheet("입고기록").get_all_records())
                df_qty = pd.DataFrame(sh.worksheet("발주기록").get_all_records())
                
                # 컬럼 공백 제거
                df_in.columns = [c.strip() for c in df_in.columns]
                df_qty.columns = [c.strip() for c in df_qty.columns]
                
                st.session_state.df_in_all = df_in
                st.session_state.df_qty_all = df_qty
                st.success("✅ 실시간 장부 동기화 완료!")
                st.rerun()
            except Exception as e:
                st.error(f"데이터 로드 실패: {e}")

    if 'df_qty_all' in st.session_state and 'df_in_all' in st.session_state:
        # 데이터 복사
        in_all = st.session_state.df_in_all.copy()
        qty_all = st.session_state.df_qty_all.copy()
        
        # 수량 컬럼 찾기 (유연한 대응)
        in_num_col = next((c for c in ['입고수량', '수량', '입고'] if c in in_all.columns), None)
        qty_num_col = next((c for c in ['발주수량', '수량', '발주'] if c in qty_all.columns), None)

        # 기준 키 (업체, 상품명, 옵션)
        base_keys = ['업체명', '상품명', '옵션', '공급처상품명']
        
        # 1. 발주 합계 계산
        qty_sum = qty_all.groupby(base_keys).agg({
            qty_num_col: 'sum',
            '날짜': 'max',
            '메모': lambda x: " / ".join(dict.fromkeys([str(i) for i in x if str(i).strip()]))
        }).reset_index()

        # 2. 입고 합계 계산
        in_sum = in_all.groupby(base_keys)[in_num_col].sum().reset_index()

        # 3. 데이터 병합 (잔량 계산)
        summary = pd.merge(qty_sum, in_sum, on=base_keys, how='left').fillna(0)
        summary['리오더 잔량'] = summary[qty_num_col] - summary[in_num_col]
        
        # 잔량 있는 것만 필터링
        summary = summary[summary['리오더 잔량'] > 0].sort_values('리오더 잔량', ascending=False)

        # 컬럼명 정리 및 순서 배치 (사장님 요청 순서)
        summary.rename(columns={'날짜': '최근기록일', qty_num_col: '총발주량', in_num_col: '입고수량', '메모': '비고(처리내역)'}, inplace=True)
        summary['발주수량'] = summary['총발주량'] # 요청하신 셀 구성을 위해 복사
        
        order = ['최근기록일', '업체명', '상품명', '옵션', '공급처상품명', '총발주량', '입고수량', '발주수량', '리오더 잔량', '비고(처리내역)']
        final_6 = summary[[c for c in order if c in summary.columns]]

        # 엑셀 다운로드
        import io
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            final_6.to_excel(writer, index=False)
        st.download_button("📥 현재 잔량 현황 엑셀 다운로드", output.getvalue(), "reorder_status.xlsx", "application/vnd.ms-excel")

        st.dataframe(final_6, use_container_width=True, hide_index=True)
    else:
        st.info("💡 위 업데이트 버튼을 눌러 입고/발주 시트 데이터를 불러오세요.")
