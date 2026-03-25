import streamlit as st

import pandas as pd

from datetime import datetime

import gspread

from oauth2client.service_account import ServiceAccountCredentials



# --- [1. 공통 함수 정의] ---

def get_sheet():

    creds_dict = dict(st.secrets["gcp_service_account"])

    scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive"]

    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)

    client = gspread.authorize(creds)

    spreadsheet_key = "1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg"

    return client.open_by_key(spreadsheet_key)



def save_reorder_data(df):

    try:

        sheet = get_sheet().sheet1

        sheet.clear()

        sheet.update([df.columns.values.tolist()] + df.values.tolist())

    except Exception as e:

        st.error(f"구글 시트 저장 실패: {e}")



def save_history_to_gsheet(df):

    try:

        spreadsheet = get_sheet()

        try:

            hist_sheet = spreadsheet.worksheet("history")

        except:

            hist_sheet = spreadsheet.add_worksheet(title="history", rows="1000", cols="20")

            hist_sheet.append_row(["저장시간"] + df.columns.tolist())

        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        rows_to_add = [[now_str] + row for row in df.values.tolist()]

        hist_sheet.append_rows(rows_to_add)

        return True

    except Exception as e:

        st.error(f"과거 기록 저장 실패: {e}")

        return False



def load_history_from_gsheet():

    try:

        spreadsheet = get_sheet()

        hist_sheet = spreadsheet.worksheet("history")

        data = hist_sheet.get_all_records()

        return pd.DataFrame(data)

    except Exception as e:

        st.error(f"기록 불러오기 실패: {e}")

        return pd.DataFrame()



def find_idx(cols, target_keywords):

    for keyword in target_keywords:

        for i, col in enumerate(cols):

            if keyword in str(col): return i

    return 0



# --- [2. 앱 설정 및 탭 구성] ---

st.set_page_config(layout="wide", page_title="재고 관리 시스템")

st.title("📦 통합 재고 관리 시스템")



tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 관리"])



# --- [🏭 탭 1: 제작 상품 관리] ---

with tab1:

    if 'analyzed' not in st.session_state: st.session_state.analyzed = False

    

    st.subheader("📁 데이터 업로드 (제작상품)")

    if st.button("🔄 제작상품 데이터 초기화"):

        st.session_state.clear()

        st.rerun()



    uploaded_file = st.file_uploader("엑셀/CSV 파일을 선택하세요", type=['xlsx', 'xls', 'csv'], key="prod_upload")

    st.divider()



    # [수정] 파일 업로드 시 구글 시트에서 기존 리오더 수량을 찾아 합칩니다.

    if uploaded_file is not None:

        if 'df_raw' not in st.session_state or st.session_state.get('last_filename') != uploaded_file.name:

            df_new = pd.read_excel(uploaded_file)

            df_new.columns = df_new.columns.str.strip()

            df_new = df_new.loc[:, ~df_new.columns.duplicated()] 

            

            # 매핑용 임시 변수 (가용재고 등 매핑 전이라도 상품명/옵션은 찾아야 함)

            tmp_item = next((c for c in df_new.columns if '상품명' in c), df_new.columns[0])

            tmp_option = next((c for c in df_new.columns if '옵션' in c), df_new.columns[min(1, len(df_new.columns)-1)])



            try:

                sheet = get_sheet().sheet1

                gs_data = pd.DataFrame(sheet.get_all_records())

                

                if not gs_data.empty and '상품명' in gs_data.columns:

                    gs_subset = gs_data[['상품명', '옵션', '리오더 수량']].copy()

                    gs_subset['상품명'] = gs_subset['상품명'].astype(str).str.strip()

                    gs_subset['옵션'] = gs_subset['옵션'].astype(str).str.strip()

                    

                    df_new['match_name'] = df_new[tmp_item].astype(str).str.strip()

                    df_new['match_opt'] = df_new[tmp_option].astype(str).str.strip()



                    df_new = pd.merge(df_new, gs_subset, left_on=['match_name', 'match_opt'], right_on=['상품명', '옵션'], how='left', suffixes=('', '_gs'))

                    

                    if '리오더 수량_gs' in df_new.columns:

                        df_new['리오더 수량'] = df_new['리오더 수량_gs'].fillna(0).astype(int)

                        # 중복 컬럼 정리

                        cols_to_drop = [c for c in ['상품명_gs', '옵션_gs', '리오더 수량_gs', 'match_name', 'match_opt'] if c in df_new.columns]

                        df_new = df_new.drop(columns=cols_to_drop)

                

                if '리오더 수량' not in df_new.columns: df_new['리오더 수량'] = 0

            except Exception as e:

                st.warning(f"기존 리오더 수량 불러오기 실패: {e}")

                if '리오더 수량' not in df_new.columns: df_new['리오더 수량'] = 0



            st.session_state.df_raw = df_new

            st.session_state.last_filename = uploaded_file.name

            st.session_state.analyzed = False

            st.rerun()



    if st.session_state.get('df_raw') is not None:

        df_current = st.session_state.df_raw

        cols = df_current.columns.tolist()



        st.subheader("⚙️ 1단계: 매핑 설정")

        c1, c2 = st.columns(2)

        sold_out = c1.selectbox("품절 여부", cols, index=find_idx(cols, ['품절']))

        vendor = c1.selectbox("공급처", cols, index=find_idx(cols, ['공급처']))

        item = c1.selectbox("상품명", cols, index=find_idx(cols, ['상품명']))

        option = c1.selectbox("옵션", cols, index=find_idx(cols, ['옵션']))

        vendor_item = c1.selectbox("공급처 상품명", cols, index=find_idx(cols, ['공급처상품명']))

        reg_date = c2.selectbox("등록일", cols, index=find_idx(cols, ['등록일']))

        stock = c2.selectbox("정상재고", cols, index=find_idx(cols, ['정상재고']))

        avail = c2.selectbox("가용재고", cols, index=find_idx(cols, ['가용재고']))

        t3day = c2.selectbox("3일 발주합계", cols, index=find_idx(cols, ['3일']))

        t1week = c2.selectbox("7일 발주합계", cols, index=find_idx(cols, ['7일', '1주']))



        st.subheader("⚙️ 2~3단계: 분석 설정")

        col_lt, col_ss = st.columns(2)

        lead_time = col_lt.number_input("리드타임 (일)", value=10)

        safety_stock = col_ss.number_input("안전재고 (일 수)", value=7)

        

        if st.button("🚀 분석 실행"):

            df = st.session_state.df_raw.copy()

            daily_avg = pd.to_numeric(df[t1week], errors='coerce').fillna(0) / 7

            df['권장 발주량'] = ((daily_avg * lead_time) + (daily_avg * safety_stock) - pd.to_numeric(df[avail], errors='coerce').fillna(0)).clip(lower=0).astype(int)

            if '리오더 수량' not in df.columns: df['리오더 수량'] = 0

            if '리오더입고수량' not in df.columns: df['리오더입고수량'] = 0

            st.session_state.df_raw = df

            st.session_state.analyzed = True

            st.rerun()



        if st.session_state.analyzed:

         # --- [4단계: 데이터 편집 및 재고 관리 - 리오더 차감 전용] ---
st.divider()
st.subheader("📊 4단계: 데이터 편집 및 재고 관리")

# 1. 데이터 복사 및 수치형 변환
df_work = st.session_state.df_raw.copy()

num_cols = [stock, avail, "리오더 수량", t7day, t3day]
for c in num_cols:
    if c in df_work.columns:
        df_work[c] = pd.to_numeric(df_work[c], errors='coerce').fillna(0).astype(int)

# 2. [계산식] 일판매량 반올림 및 권장발주량
v7 = df_work[t7day]
v3 = df_work[t3day]

# 💡 일판매량: 반올림 후 정수 처리 (소수점 제거)
df_work['일판매량'] = (v7 / 7 if v7.sum() > 0 else v3 / 3).round(0).astype(int)

# 권장발주량 계산
df_work['권장발주량'] = ((df_work['일판매량'] * (lt + ss)) - (df_work[avail] + df_work['리오더 수량'])).clip(lower=0).astype(int)
df_work['3일발주합계'] = df_work[t3day]

# 3. 상단 UI 및 필터
f_c1, f_c2, f_c3 = st.columns([2, 1, 1])
search_q = f_c1.text_input("🔍 상품명 검색", key="search_v4_input_final_v2")
filter_m = f_c2.selectbox("품절 필터", ["전체보기", "정상만", "품절만"], index=1, key="filter_v4_select_final_v2")
hist_date_4 = f_c3.date_input("🗓️ 입고 매핑 날짜", datetime.now(KST).date(), key="date_v4_input_final_v2")

if filter_m == "정상만": df_work = df_work[~df_work[sold_out].astype(str).str.contains('품절', na=False)]
elif filter_m == "품절만": df_work = df_work[df_work[sold_out].astype(str).str.contains('품절', na=False)]
if search_q: df_work = df_work[df_work[item].astype(str).str.contains(search_q, case=False, na=False)]

# 🎯 [순서 및 명칭 정리] 사장님 요청 13개 컬럼
df_display = df_work.rename(columns={
    sold_out: "품절", vendor: "공급쳐", v_item: "공급쳐 상품명",
    item: "상품명", option: "옵션", stock: "정상재고", avail: "가용재고",
    "리오더입고수량": "리오더 입고수량", "과거 리오더입고": "과거리오더 입고"
})

final_cols = [
    "품절", "공급쳐", "상품명", "옵션", "공급쳐 상품명", 
    "정상재고", "가용재고", "리오더 수량", "리오더 입고수량", 
    "과거리오더 입고", "3일발주합계", "일판매량", "권장발주량"
]
actual_final_cols = [c for c in final_cols if c in df_display.columns]

# 4. 저장 폼 및 차감 로직
with st.form("form_step_4_reorder_only_fix"):
    edited_v4 = st.data_editor(df_display[actual_final_cols], use_container_width=True, key="editor_v4_reorder_fix", hide_index=True)
    submit_v4 = st.form_submit_button("💾 입고량 반영 및 저장", use_container_width=True, type="primary")
    
    if submit_v4:
        # 💡 [로딩 액션 시작] 여기서부터 아래 작업이 끝날 때까지 로딩바가 돕니다.
        with st.spinner('📡 입고 데이터를 기록하고 리오더 수량을 차감 중입니다...'):
            edits = st.session_state["editor_v4_reorder_fix"].get("edited_rows", {})
            if edits:
                for r_idx_str, change in edits.items():
                    orig_idx = df_work.index[int(r_idx_str)]
                    
                    # 1) 리오더 수량 직접 수정 시 반영
                    if "리오더 수량" in change:
                        st.session_state.df_raw.at[orig_idx, "리오더 수량"] = int(change["리오더 수량"])
                    
                    # 2) 리오더 입고수량 입력 시 -> 차감 로직
                    if "리오더 입고수량" in change:
                        in_qty = int(change["리오더 입고수량"])
                        if in_qty > 0:
                            current_reorder = int(st.session_state.df_raw.at[orig_idx, "리오더 수량"])
                            st.session_state.df_raw.at[orig_idx, "리오더 수량"] = max(0, current_reorder - in_qty)
                            
                            # 입고 히스토리 저장
                            log_df = pd.DataFrame([[df_work.at[orig_idx, item], df_work.at[orig_idx, option], in_qty]], columns=['상품명', '옵션', '수량'])
                            save_history_to_gsheet(log_df, log_type="입고")

                # 최종 저장 및 화면 갱신
                save_reorder_data(st.session_state.df_raw[[item, option, '리오더 수량']].rename(columns={item:'상품명', option:'옵션'}))
                st.success("✅ 리오더 수량 차감 및 저장이 완료되었습니다!")
                time.sleep(1) # 사장님이 성공 메시지를 보실 수 있게 1초 대기
                st.rerun()


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
