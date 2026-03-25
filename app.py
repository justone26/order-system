import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

# --- [0. 상수 설정] ---
KST = timezone(timedelta(hours=9)) # 한국 시간대 설정

# --- [1. 공통 함수 정의] ---
def get_sheet():
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        return client.open_by_key("1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg")
    except Exception as e:
        st.error(f"구글 시트 연결 실패: {e}")
        return None

def save_reorder_data(df, item_col, opt_col):
    """현재 세션의 리오더 수량을 구글 시트 메인 탭에 덮어쓰기 (보존용)"""
    try:
        spreadsheet = get_sheet()
        if spreadsheet:
            sheet = spreadsheet.sheet1
            sheet.clear()
            # 상품명, 옵션, 리오더 수량 컬럼만 추출하여 저장
            save_df = df[[item_col, opt_col, '리오더 수량']].copy()
            save_df.columns = ['상품명', '옵션', '리오더 수량']
            sheet.update([save_df.columns.values.tolist()] + save_df.values.tolist())
    except Exception as e:
        st.error(f"데이터 보존 실패: {e}")

# --- [2. 앱 설정] ---
st.set_page_config(layout="wide", page_title="저스트원 재고관리")
st.title("📦 저스트원 통합 재고 관리 시스템")

tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 관리"])

with tab1:
    if 'df_raw' not in st.session_state: st.session_state.df_raw = None
    if 'analyzed' not in st.session_state: st.session_state.analyzed = False

    st.subheader("📁 1단계: 데이터 업로드")
    c_up1, c_up2 = st.columns([3, 1])
    uploaded_file = c_up1.file_uploader("엑셀/CSV 파일을 선택하세요", type=['xlsx', 'xls', 'csv'], key="main_up")
    
    if c_up2.button("🔄 전체 데이터 초기화", use_container_width=True):
        st.session_state.clear()
        st.rerun()

    # --- [핵심] 업체별 데이터 누적 및 리오더 보존 로직 ---
    if uploaded_file is not None:
        if st.session_state.get('last_fn') != uploaded_file.name:
            with st.spinner(f'{uploaded_file.name} 데이터를 처리 중입니다...'):
                # 1. 새 파일 읽기
                df_new = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
                df_new.columns = df_new.columns.str.strip()
                df_new = df_new.loc[:, ~df_new.columns.duplicated()]

                # 2. 리오더 수량 매칭 (구글 시트 연동)
                try:
                    gs = get_sheet()
                    gs_data = pd.DataFrame(gs.sheet1.get_all_records())
                    if not gs_data.empty and '리오더 수량' in gs_data.columns:
                        tmp_item = next((c for c in df_new.columns if '상품명' in c), df_new.columns[0])
                        tmp_opt = next((c for c in df_new.columns if '옵션' in c), df_new.columns[1])
                        
                        df_new['tmp_n'] = df_new[tmp_item].astype(str).str.strip()
                        df_new['tmp_o'] = df_new[tmp_opt].astype(str).str.strip()
                        gs_data['상품명'] = gs_data['상품명'].astype(str).str.strip()
                        gs_data['옵션'] = gs_data['옵션'].astype(str).str.strip()
                        
                        df_new = pd.merge(df_new, gs_data[['상품명', '옵션', '리오더 수량']], 
                                         left_on=['tmp_n', 'tmp_o'], right_on=['상품명', '옵션'], 
                                         how='left', suffixes=('', '_gs'))
                        
                        if '리오더 수량_gs' in df_new.columns:
                            df_new['리오더 수량'] = df_new['리오더 수량_gs'].fillna(0).astype(int)
                            df_new.drop(columns=['상품명_gs', '옵션_gs', '리오더 수량_gs', 'tmp_n', 'tmp_o'], inplace=True, errors='ignore')
                except:
                    if '리오더 수량' not in df_new.columns: df_new['리오더 수량'] = 0

                # 3. [데이터 누적] 기존 데이터가 있으면 합치기
                if st.session_state.df_raw is not None:
                    # 기존 데이터 + 새 데이터 합치기
                    st.session_state.df_raw = pd.concat([st.session_state.df_raw, df_new], ignore_index=True)
                    # 중복 제거 (상품명/옵션 기준)
                    target_n = next((c for c in st.session_state.df_raw.columns if '상품명' in c), st.session_state.df_raw.columns[0])
                    target_o = next((c for c in st.session_state.df_raw.columns if '옵션' in c), st.session_state.df_raw.columns[1])
                    st.session_state.df_raw.drop_duplicates(subset=[target_n, target_o], keep='last', inplace=True)
                    st.success(f"✅ 기존 데이터에 {uploaded_file.name} 업체가 추가되었습니다!")
                else:
                    st.session_state.df_raw = df_new
                    st.success(f"✅ {uploaded_file.name} 데이터가 로드되었습니다.")

                st.session_state.last_fn = uploaded_file.name
                st.session_state.analyzed = False
                st.rerun()

    # --- [매핑 및 분석 로직] ---
    if st.session_state.df_raw is not None:
        df_curr = st.session_state.df_raw
        cols = df_curr.columns.tolist()

        st.divider()
        st.subheader("⚙️ 2단계: 매핑 설정")
        c1, c2 = st.columns(2)
        with c1:
            sold_out = st.selectbox("품절 여부", cols, index=0)
            vendor = st.selectbox("공급처", cols, index=0)
            v_item = st.selectbox("공급처 상품명", cols, index=0)
            item = st.selectbox("상품명", cols, index=0)
            option = st.selectbox("옵션", cols, index=0)
        with c2:
            reg_date = st.selectbox("등록일", cols, index=0)
            stock = st.selectbox("정상재고", cols, index=0)
            avail = st.selectbox("가용재고", cols, index=0)
            t3day = st.selectbox("3일 발주합계", cols, index=0)
            t7day = st.selectbox("7일 발주합계", cols, index=0)

        if st.button("🚀 분석 실행", use_container_width=True, type="primary"):
            st.session_state.analyzed = True
            st.rerun()

        if st.session_state.analyzed:
            st.divider()
    
# --- [결과 화면 및 4단계 로직 시작] ---
        if st.session_state.analyzed and st.session_state.df_raw is not None:
            st.divider()
            st.subheader("📊 4단계: 데이터 편집 및 재고 관리")

            # 1. 데이터 복사 및 수치형 변환
            df_work = st.session_state.df_raw.copy()

            # 수치 데이터 컬럼 변환 (에러 방지)
            num_cols = [stock, avail, "리오더 수량", t7day, t3day]
            for c in num_cols:
                if c in df_work.columns:
                    df_work[c] = pd.to_numeric(df_work[c], errors='coerce').fillna(0).astype(int)

            # 리오더 입고수량 컬럼이 없으면 생성
            if "리오더 입고수량" not in df_work.columns:
                df_work["리오더 입고수량"] = 0

            # 2. [계산식] 일판매량 및 권장발주량 (lt=리드타임, ss=안전재고)
            lt = 3; ss = 2 
            v7 = df_work[t7day]
            v3 = df_work[t3day]

            # 일판매량: 7일 우선, 데이터 없으면 3일 기준 반올림
            df_work['일판매량'] = (v7 / 7 if v7.sum() > 0 else v3 / 3).round(0).astype(int)

            # 권장발주량 계산: (일판매량 * (리드타임+안전재고)) - (가용재고 + 리오더수량)
            df_work['권장발주량'] = ((df_work['일판매량'] * (lt + ss)) - (df_work[avail] + df_work['리오더 수량'])).clip(lower=0).astype(int)
            df_work['3일발주합계'] = df_work[t3day]

            # 3. 상단 UI 및 필터
            f_c1, f_c2, f_c3 = st.columns([2, 1, 1])
            search_q = f_c1.text_input("🔍 상품명 검색", key="search_v4_input_final_v2")
            filter_m = f_c2.selectbox("품절 필터", ["전체보기", "정상만", "품절만"], index=1, key="filter_v4_select_final_v2")
            hist_date_4 = f_c3.date_input("🗓️ 입고 매핑 날짜", datetime.now(KST).date(), key="date_v4_input_final_v2")

            # 필터 적용 로직
            if filter_m == "정상만": 
                df_work = df_work[~df_work[sold_out].astype(str).str.contains('품절', na=False)]
            elif filter_m == "품절만": 
                df_work = df_work[df_work[sold_out].astype(str).str.contains('품절', na=False)]
            if search_q: 
                df_work = df_work[df_work[item].astype(str).str.contains(search_q, case=False, na=False)]

            # 🎯 컬럼명 정리
            df_display = df_work.rename(columns={
                sold_out: "품절", vendor: "공급쳐", v_item: "공급쳐 상품명",
                item: "상품명", option: "옵션", stock: "정상재고", avail: "가용재고"
            })

            final_cols = [
                "품절", "공급쳐", "상품명", "옵션", "공급쳐 상품명", 
                "정상재고", "가용재고", "리오더 수량", "리오더 입고수량", 
                "3일발주합계", "일판매량", "권장발주량"
            ]
            actual_final_cols = [c for c in final_cols

# --- [5단계: 최종 발주 리스트 요약 - 저장 및 엑셀 버튼 복구] ---
st.divider()
st.subheader("📋 5단계: 최종 발주 리스트 요약")

if 'add_order_dict' not in st.session_state: 
    st.session_state.add_order_dict = {}

df_5 = st.session_state.df_raw.copy()

# 숫자형 변환 및 일판매량(반올림) 계산
num_cols_5 = [avail, '리오더 수량', t7day, t3day]
for c in num_cols_5:
    if c in df_5.columns:
        df_5[c] = pd.to_numeric(df_5[c], errors='coerce').fillna(0).astype(int)

v7_5 = df_5[t7day]; v3_5 = df_5[t3day]
df_5['일판매량'] = (v7_5 / 7 if v7_5.sum() > 0 else v3_5 / 3).round(0).astype(int)
df_5['권장발주량'] = ((df_5['일판매량'] * (lt + ss)) - (df_5[avail] + df_5['리오더 수량'])).clip(lower=0).astype(int)
df_5['추가발주수량'] = df_5.index.map(st.session_state.add_order_dict).fillna(0).astype(int)

# 상태 판별
def get_final_status(r):
    stock_sum = r[avail] + r['리오더 수량']; daily = r['일판매량']
    if daily > 0:
        if stock_sum < (daily * 3): return "🚨 긴급"
        if stock_sum < (daily * 5): return "⚠️ 주의"
    return "✅ 정상"
df_5['상태'] = df_5.apply(get_final_status, axis=1)

# 검색 및 필터 UI
c5_1, c5_2, c5_3 = st.columns([1.5, 1.5, 1])
search_q_v5 = c5_2.text_input("🔍 전체 상품명 검색", key="v5_ordered_final_fix")
s_filter = c5_1.selectbox("🎯 상태 필터", ["🚨긴급 + ⚠️주의 우선", "🚨 긴급만 보기", "✅ 전체보기"], index=0)
hist_date_5 = c5_3.date_input("🗓️ 기록 확인 날짜", datetime.now(KST).date()) # KST 적용

if search_q_v5:
    df_5 = df_5[df_5[item].astype(str).str.contains(search_q_v5, case=False, na=False)]
else:
    if s_filter == "🚨긴급 + ⚠️주의 우선": 
        df_5 = df_5[df_5['상태'].isin(["🚨 긴급", "⚠️ 주의"]) | (df_5['권장발주량'] > 0)]
    elif s_filter == "🚨 긴급만 보기": 
        df_5 = df_5[df_5['상태'] == "🚨 긴급"]

df_5 = df_5.sort_values(by='상태')

# 🎯 [순서 정리]
df_display_5 = df_5.rename(columns={item: "상품명", option: "옵션", v_item: "공급쳐상품명", avail: "가용재고", "리오더 수량": "리오더수량"})
final_cols_5 = ["상태", "상품명", "옵션", "공급쳐상품명", "가용재고", "리오더수량", "추가발주수량", "권장발주량"]
actual_cols_5 = [c for c in final_cols_5 if c in df_display_5.columns]

# 4. 데이터 에디터
with st.form("form_step_5_final_v15"):
    edited_v5 = st.data_editor(df_display_5[actual_cols_5], use_container_width=True, key="editor_v5_v15", hide_index=True)
    
    # 💡 로딩 액션 추가 (수량 확정 버튼)
    if st.form_submit_button("✅ 수량 확정 (리오더 수량 합산)", use_container_width=True, type="primary"):
        with st.spinner('🔄 리오더 수량을 합산하여 갱신 중입니다...'):
            edits = st.session_state["editor_v5_v15"].get("edited_rows", {})
            if edits:
                for r_idx_str, change in edits.items():
                    orig_idx = df_5.index[int(r_idx_str)]
                    if "추가발주수량" in change:
                        add_qty = int(change["추가발주수량"])
                        st.session_state.df_raw.at[orig_idx, "리오더 수량"] += add_qty
                        st.session_state.add_order_dict[orig_idx] = add_qty
                st.success("✅ 리오더 수량이 갱신되었습니다.")
                time.sleep(1)
                st.rerun()

# --- [5단계 하단: 저장 및 엑셀 버튼] ---
st.write("---")
b1, b2 = st.columns(2)

# 1. 구글 시트 저장 버튼 (로딩 액션 추가)
if b1.button("💾 구글 시트에 최종 발주 기록 저장", use_container_width=True):
    with st.spinner('📡 한국 시간으로 발주 데이터를 구글 시트에 안전하게 저장 중입니다...'):
        order_ready = df_5[(df_5['권장발주량'] > 0) | (df_5['추가발주수량'] > 0)].copy()
        
        if not order_ready.empty:
            now_kst = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
            order_ready['저장시간'] = now_kst
            order_ready['공급쳐상품명'] = order_ready[v_item]
            order_ready['가용재고'] = order_ready[avail]
            order_ready['리오더수량_저장'] = order_ready['리오더 수량'] 
            order_ready['추가발주수량_저장'] = order_ready['추가발주수량']
            order_ready['권장발주수량_저장'] = order_ready['권장발주량']
            
            save_data = order_ready[[
                '저장시간', item, option, '공급쳐상품명', 
                '가용재고', '리오더수량_저장', '추가발주수량_저장', '권장발주수량_저장'
            ]]
            
            save_data.columns = [
                "저장시간", "상품명", "옵션", "공급쳐상품명", 
                "가용재고", "리오더수량", "추가발주수량", "권장발주수량"
            ]
            
            if save_history_to_gsheet(save_data, log_type="발주"):
                st.success(f"✅ 한국 시간({now_kst})으로 모든 데이터가 저장되었습니다!")
                time.sleep(1)
                st.rerun()
        else:
            st.warning("발주할 수량이 있는 상품이 없습니다.")

# 2. 📥 엑셀 다운로드 버튼
if not df_display_5.empty:
    csv_final = df_display_5[actual_cols_5].to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
    b2.download_button(
        label="📥 현재 리스트 엑셀 다운로드",
        data=csv_final,
        file_name=f"최종발주서_{datetime.now(KST).strftime('%m%d_%H%M')}.csv",
        mime="text/csv",
        use_container_width=True
    )


# --- [6단계: 전체 히스토리 내역 - 한국 시간 및 8대 항목 완결판] ---
st.divider()
st.subheader("📜 6단계: 전체 히스토리 내역")

# 1. 구글 시트에서 전체 데이터 로드
with st.spinner('📡 히스토리 기록을 불러오는 중입니다...'):
    df_hist = load_history_from_gsheet()

if not df_hist.empty:
    # 💡 [명칭 정리] 시트의 컬럼명을 화면 표시용으로 통일
    rename_map = {
        'Log Type': '구분',
        'v_item': '공급쳐상품명',
        '리오더수량_저장': '리오더수량',
        '추가발주수량_저장': '추가발주수량',
        '권장발주수량_저장': '권장발주수량'
    }
    df_hist = df_hist.rename(columns=rename_map)

    # 2. 상단 UI (달력 및 검색창)
    h_c1, h_c2 = st.columns([1, 2])
    
    # 달력 기본값을 한국 시간(KST) 오늘 날짜로 설정
    today_kst = datetime.now(KST).date()
    h_date = h_c1.date_input("🗓️ 조회 날짜 선택", today_kst, key="h_date_v6_final")
    h_search = h_c2.text_input("🔍 상품명 검색", key="h_search_v6_final")

    # 3. 데이터 필터링
    # 저장시간 컬럼에서 날짜만 추출하여 달력과 비교
    if '저장시간' in df_hist.columns:
        df_hist['날짜_tmp'] = pd.to_datetime(df_hist['저장시간']).dt.date
        df_hist = df_hist[df_hist['날짜_tmp'] == h_date]
    
    # 검색어 필터링 (상품명 기준)
    if h_search:
        df_hist = df_hist[df_hist['상품명'].astype(str).str.contains(h_search, case=False, na=False)]

    # 🎯 [순서 고정] 사장님이 요청하신 8가지 항목
    # 저장시간, 상품명, 옵션, 공급쳐상품명, 가용재고, 리오더수량, 추가발주수량, 권장발주수량
    final_view_cols = [
        "저장시간", "상품명", "옵션", "공급쳐상품명", 
        "가용재고", "리오더수량", "추가발주수량", "권장발주수량"
    ]
    
    # 실제 데이터에 존재하는 컬럼만 선별 (에러 방지)
    actual_view = [c for c in final_view_cols if c in df_hist.columns]

    # 4. 결과 출력
    if not df_hist.empty:
        # 최신 저장 시간이 위로 오도록 정렬하여 출력
        st.dataframe(
            df_hist[actual_view].sort_values(by='저장시간', ascending=False), 
            use_container_width=True, 
            hide_index=True
        )
        
        # 5. 📥 엑셀 다운로드 버튼 (현재 화면에 보이는 8개 항목 그대로)
        csv_hist = df_hist[actual_view].to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
        st.download_button(
            label=f"📥 {h_date} 발주 내역 엑셀 다운로드",
            data=csv_hist,
            file_name=f"발주히스토리_{h_date}.csv",
            mime="text/csv",
            use_container_width=True
        )
    else:
        st.info(f"📅 {h_date} 날짜에는 저장된 기록이 없습니다. 다른 날짜를 선택하거나 5단계에서 저장해 주세요.")
else:
    st.warning("아직 저장된 히스토리 기록이 없습니다. 5단계에서 '기록 저장' 버튼을 먼저 눌러주세요.")



# --- [🌙 탭 2: 동대문 사입 관리] ---

with tab2:

    st.subheader("🌙 동대문 사입 및 미납 관리")

    dong_file = st.file_uploader("동대문 주문 리스트 업로드", type=['xlsx', 'csv'], key="dong_tab_upload")

    if dong_file:

        if "last_file_name" not in st.session_state or st.session_state.last_file_name != dong_file.name:

            df = pd.read_excel(dong_file)

            df.columns = df.columns.str.strip()

            required_cols = ['선택', '품절', '상품명', '공급처', '공급처상품명', '정상재고', '가용재고', '판매수량', '발주수량', '가중율', '3일판매']

            for col in required_cols:

                if col not in df.columns: df[col] = 0 if col not in ['선택', '품절', '상품명', '공급처', '공급처상품명'] else ""

            for col in ['정상재고', '가용재고', '3일판매']: df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            df['판매수량'] = (df['정상재고'] - df['가용재고']).clip(lower=0)

            df['가중율'] = df['판매수량'].apply(lambda n: 2.0 if n >= 10 else (1.5 if n >= 6 else (1.2 if n >= 3 else 1.0)))

            df['발주수량'] = (df['판매수량'] * df['가중율']).astype(int)

            st.session_state.df_dong_current = df[required_cols]

            st.session_state.last_file_name = dong_file.name



        df_display = st.session_state.df_dong_current.copy()

        search_query = st.text_input("상품명 검색 (사입)")

        if search_query: df_display = df_display[df_display['상품명'].astype(str).str.contains(search_query, case=False, na=False)]

        

        df_display['선택'] = df_display['선택'].astype(bool)

        edited_df = st.data_editor(df_display, use_container_width=True, key="dong_editor")

        

        st.divider()

        c1, c2, c3 = st.columns(3)

        add_val = c1.number_input("추가 수량", value=1, min_value=1)

        if c2.button("🚀 선택 상품 수량 더하기"):

            selected = edited_df[edited_df['선택'] == True].index

            for idx in selected: st.session_state.df_dong_current.at[idx, '발주수량'] += add_val

            st.rerun()

        csv = edited_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')

        c3.download_button("📥 엑셀 다운로드", csv, "사입리스트.csv", "text/csv")
