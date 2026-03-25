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
            
            # 구글 시트 연동 로직
            try:
                sheet = get_sheet().sheet1
                gs_data = pd.DataFrame(sheet.get_all_records())
                
                if not gs_data.empty and '상품명' in gs_data.columns:
                    # 매칭을 위해 상품명, 옵션, 리오더 수량만 가져옴
                    gs_subset = gs_data[['상품명', '옵션', '리오더 수량']].copy()
                    # 기존 데이터와 합치기 (상품명+옵션 기준)
                    df_new = pd.merge(df_new, gs_subset, on=['상품명', '옵션'], how='left', suffixes=('', '_gs'))
                    
                    if '리오더 수량_gs' in df_new.columns:
                        df_new['리오더 수량'] = df_new['리오더 수량_gs'].fillna(0)
                        df_new = df_new.drop(columns=['리오더 수량_gs'])
                else:
                    df_new['리오더 수량'] = 0
            except Exception as e:
                st.warning(f"기존 리오더 수량 불러오기 실패: {e}")
                df_new['리오더 수량'] = 0

            st.session_state.df_raw = df_new
            st.session_state.last_filename = uploaded_file.name
            st.session_state.analyzed = False
            st.rerun()

    # 데이터가 로드된 경우에만 매핑 및 분석 진행
    if st.session_state.get('df_raw') is not None:
        df_current = st.session_state.df_raw
        cols = df_current.columns.tolist()

        # 1단계: 매핑 설정 (여기서 NameError 방지를 위해 find_idx가 코드 상단에 있어야 함)
        st.subheader("⚙️ 1단계: 매핑 설정")
        c1, c2 = st.columns(2)
        
        # find_idx 함수는 코드 맨 위에 정의되어 있어야 에러가 안 납니다.
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

        # 2~3단계: 분석 설정
        st.subheader("⚙️ 2~3단계: 분석 설정")
        col_lt, col_ss = st.columns(2)
        lead_time = col_lt.number_input("리드타임 (일)", value=10)
        safety_stock = col_ss.number_input("안전재고 (일 수)", value=7)
        
        if st.button("🚀 분석 실행"):
            df = st.session_state.df_raw.copy()
            # 판매량 계산 및 발주량 계산
            daily_avg = pd.to_numeric(df[t1week], errors='coerce').fillna(0) / 7
            df['권장 발주량'] = ((daily_avg * lead_time) + (daily_avg * safety_stock) - pd.to_numeric(df[avail], errors='coerce').fillna(0)).clip(lower=0).astype(int)
            
            # 리오더 수량 컬럼이 없으면 생성
            if '리오더 수량' not in df.columns: df['리오더 수량'] = 0
            if '리오더입고수량' not in df.columns: df['리오더입고수량'] = 0
            
            st.session_state.df_raw = df
            st.session_state.analyzed = True
            st.rerun()


if st.session_state.get('analyzed') and st.session_state.get('p'):
    # --- [4단계: 데이터 편집 및 재고 관리] ---
    st.divider()
    st.subheader("📊 4단계: 데이터 편집 및 재고 관리")

    p = st.session_state.p
    sold_out, vendor, v_item, item, option = p['so'], p['vn'], p['vi'], p['it'], p['op']
    stock, avail, t3day, t7day = p['st'], p['av'], p['t3'], p['t7']
    lt, ss = p['lt'], p['ss']

    df_work = st.session_state.df_raw.copy()
    
    # 입고 기록 실시간 매칭
    incoming_hist = get_incoming_history()
    if not incoming_hist.empty:
        df_work = pd.merge(df_work, incoming_hist, left_on=[item, option], right_on=['상품명', '옵션'], how='left')
        df_work['과거리오더 입고'] = df_work['과거리오더 입고'].fillna(0).astype(int)
    else:
        df_work['과거리오더 입고'] = 0

    # 숫자형 변환
    for c in [stock, avail, t7day, t3day]:
        if c in df_work.columns:
            df_work[c] = pd.to_numeric(df_work[c], errors='coerce').fillna(0).astype(int)
    
    if "리오더 수량" not in df_work.columns: df_work["리오더 수량"] = 0
    df_work["리오더 수량"] = pd.to_numeric(df_work["리오더 수량"], errors='coerce').fillna(0).astype(int)
    if "리오더 입고수량" not in df_work.columns: df_work["리오더 입고수량"] = 0

    df_work['일판매량'] = df_work.apply(lambda x: round(x[t7day] / 7) if x[t7day] > 0 else round(x[t3day] / 3), axis=1).astype(int)
    df_work['권장발주량'] = ((df_work['일판매량'] * (lt + ss)) - (df_work[avail] + df_work['리오더 수량'])).clip(lower=0).astype(int)

    # UI 및 필터
    f_c1, f_c2, f_c3 = st.columns([2, 1, 1])
    search_q = f_c1.text_input("🔍 상품명/옵션 검색", key="v4_fin_s")
    filter_m = f_c2.selectbox("상태 필터", ["전체보기", "정상만", "품절만"], index=1, key="v4_fin_f")
    hist_date_4 = f_c3.date_input("🗓️ 입고 기록 날짜", datetime.now(KST).date(), key="v4_fin_d")

    if filter_m == "정상만":
        df_work = df_work[~df_work[sold_out].astype(str).str.contains('품절', na=False)]
    elif filter_m == "품절만":
        df_work = df_work[df_work[sold_out].astype(str).str.contains('품절', na=False)]
    
    if search_q:
        df_work = df_work[df_work[item].astype(str).str.contains(search_q, case=False, na=False) | 
                          df_work[option].astype(str).str.contains(search_q, case=False, na=False)]

    df_display = df_work.rename(columns={sold_out: "품절", vendor: "공급쳐", v_item: "공급쳐 상품명", item: "상품명", option: "옵션", stock: "정상재고", avail: "가용재고"})
    final_cols = ["품절", "공급쳐", "상품명", "옵션", "공급쳐 상품명", "정상재고", "가용재고", "리오더 수량", "리오더 입고수량", "과거리오더 입고", "일판매량", "권장발주량"]

    # --- 에러 해결 포인트: alignment="left" 삭제 ---
    with st.form("form_v4_safe"):
        edited_v4 = st.data_editor(
            df_display[final_cols], 
            use_container_width=True, 
            key="ed_v4_safe", 
            hide_index=True,
            column_config={
                "정상재고": st.column_config.NumberColumn(format="%d"),
                "가용재고": st.column_config.NumberColumn(format="%d"),
                "리오더 수량": st.column_config.NumberColumn(format="%d"),
                "리오더 입고수량": st.column_config.NumberColumn(format="%d"),
                "과거리오더 입고": st.column_config.NumberColumn(format="%d"),
                "일판매량": st.column_config.NumberColumn(format="%d"),
                "권장발주량": st.column_config.NumberColumn(format="%d")
            }
        )
        if st.form_submit_button("💾 입고량 반영 및 저장", use_container_width=True, type="primary"):
            edits = st.session_state["ed_v4_safe"].get("edited_rows", {})
            if edits:
                for r_idx, change in edits.items():
                    orig_idx = df_work.index[int(r_idx)]
                    if "리오더 수량" in change:
                        st.session_state.df_raw.at[orig_idx, "리오더 수량"] = int(change["리오더 수량"])
                    if "리오더 입고수량" in change:
                        in_qty = int(change["리오더 입고수량"])
                        if in_qty > 0:
                            curr = int(st.session_state.df_raw.at[orig_idx, "리오더 수량"])
                            st.session_state.df_raw.at[orig_idx, "리오더 수량"] = max(0, curr - in_qty)
                            log_df = pd.DataFrame([[f"{hist_date_4}", df_work.at[orig_idx, item], df_work.at[orig_idx, option], in_qty]], columns=['날짜', '상품명', '옵션', '수량'])
                            save_history_to_gsheet(log_df, log_type="입고")
                save_reorder_data(st.session_state.df_raw, item, option)
                st.success("✅ 저장 완료!")
                time.sleep(1); st.rerun()

# ==========================================
# 5단계: 최종 발주 및 히스토리 자동 기록
# ==========================================

# 1. 안전 장치: 분석이 완료되었고 데이터가 있을 때만 실행
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📋 5단계: 최종 발주 리스트 요약")

    # [중요] NameError 방지를 위한 컬럼명 변수 재선언
    # 사장님 엑셀의 실제 컬럼명과 일치해야 합니다.
    avail = "가용재고"
    t7day = "7일판매량"
    t3day = "3일판매량"
    item = "상품명"
    option = "옵션"
    v_item = "공급처상품명"
    vendor = "공급처"
    lt = 7   # 리드타임 기본값
    ss = 3   # 안전재고 기본값

    # 데이터 복사 및 계산용 전처리
    df_5 = st.session_state.df_raw.copy()
    
    # 숫자 데이터 변환 (에러 방지)
    for c in [avail, '리오더 수량', t7day, t3day]:
        if c in df_5.columns:
            df_5[c] = pd.to_numeric(df_5[c], errors='coerce').fillna(0).astype(int)

    # 추가발주수량 세션 상태 확인
    if 'add_order_dict' not in st.session_state:
        st.session_state.add_order_dict = {}

    # 판매량 및 발주량 계산 로직
    df_5['일판매량'] = df_5.apply(lambda x: round(x[t7day] / 7) if x[t7day] > 0 else round(x[t3day] / 3), axis=1).astype(int)
    df_5['권장발주량'] = ((df_5['일판매량'] * (lt + ss)) - (df_5[avail] + df_5['리오더 수량'])).clip(lower=0).astype(int)
    df_5['추가발주수량'] = df_5.index.map(st.session_state.add_order_dict).fillna(0).astype(int)

    # 상태 판별 함수
    def get_stat_v5_final(r):
        tot = r[avail] + r['리오더 수량']
        day = r['일판매량']
        if day > 0:
            if tot < (day * 3): return "🚨 긴급"
            if tot < (day * 5): return "⚠️ 주의"
        return "✅ 정상"
    df_5['상태'] = df_5.apply(get_stat_v5_final, axis=1)

    # 화면 표시용 컬럼 정리
    df_disp_5 = df_5.rename(columns={item: "상품명", option: "옵션", v_item: "공급쳐상품명", avail: "가용재고", "리오더 수량": "리오더수량"})
    display_cols = ["상태", "상품명", "옵션", "공급쳐상품명", "가용재고", "리오더수량", "추가발주수량", "권장발주량"]

    # 2. 데이터 에디터 (수량 수정 가능)
    with st.form("final_order_form"):
        edited_df = st.data_editor(
            df_disp_5[display_cols],
            use_container_width=True,
            hide_index=True,
            key="v5_editor",
            column_config={
                "가용재고": st.column_config.NumberColumn(format="%d"),
                "리오더수량": st.column_config.NumberColumn(format="%d"),
                "추가발주수량": st.column_config.NumberColumn(format="%d"),
                "권장발주량": st.column_config.NumberColumn(format="%d")
            }
        )
        
        # [버튼 1] 수량 확정 및 리오더 반영
        if st.form_submit_button("✅ 수량 확정 및 리오더 반영", use_container_width=True, type="primary"):
            changes = st.session_state["v5_editor"].get("edited_rows", {})
            if changes:
                for r_idx, change in changes.items():
                    orig_idx = df_5.index[int(r_idx)]
                    if "추가발주수량" in change:
                        val = int(change["추가발주수량"])
                        st.session_state.df_raw.at[orig_idx, "리오더 수량"] += val
                        st.session_state.add_order_dict[orig_idx] = val
                # 리오더 수량 시트 저장 (이미 정의된 함수 호출)
                save_reorder_data(st.session_state.df_raw, item, option)
                st.success("✅ 리오더 수량이 업데이트되었습니다.")
                time.sleep(1); st.rerun()

    # 3. 하단 버튼 구역 (시트 저장 및 CSV 다운로드)
    st.write("---")
    col_b1, col_b2 = st.columns(2)

    with col_b1:
        # [버튼 2] 구글 시트 히스토리 저장 (6단계에서 보게 될 데이터)
        if st.button("💾 구글 시트에 최종 발주 기록 저장", use_container_width=True):
            ready = df_5.copy()
            ready['총발주'] = ready['권장발주량'] + ready['추가발주수량']
            final_to_save = ready[ready['총발주'] > 0] # 발주할 게 있는 것만 저장

            if not final_to_save.empty:
                now_str = datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')
                log_rows = []
                for _, row in final_to_save.iterrows():
                    log_rows.append([
                        now_str,                 # 날짜
                        row['상태'],              # 상태
                        row[item],               # 상품명
                        row[option],             # 옵션
                        row[v_item],             # 공급쳐상품명
                        int(row[avail]),         # 가용재고
                        int(row['리오더 수량']),   # 리오더수량
                        int(row['추가발주수량']),  # 추가발주수량
                        int(row['권장발주량'])     # 권장발주량
                    ])
                
                try:
                    sheet = get_sheet()
                    record_ws = sheet.worksheet("발주기록")
                    record_ws.append_rows(log_rows)
                    st.success(f"✅ {len(log_rows)}건의 내역이 6단계 히스토리에 저장되었습니다!")
                    st.session_state.add_order_dict = {} # 저장 후 초기화
                    time.sleep(1); st.rerun()
                except Exception as e:
                    st.error(f"📡 시트 저장 실패: {e}")
            else:
                st.warning("발주할 항목이 없습니다.")

    with col_b2:
        # [버튼 3] CSV 파일 다운로드
        csv_final = df_disp_5[display_cols].to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
        st.download_button(
            label="📥 최종 발주서 CSV 다운로드",
            data=csv_final,
            file_name=f"발주서_{datetime.now(KST).strftime('%m%d')}.csv",
            mime="text/csv",
            use_container_width=True
        )
# ==========================================
# 6단계 전용 데이터 로드 함수 (NameError 방지)
# ==========================================
def load_v6_history_final():
    try:
        # 기존에 설정된 get_sheet 함수를 사용하여 시트 접속
        sheet = get_sheet() 
        record_sheet = sheet.worksheet("발주기록")
        data = record_sheet.get_all_records()
        return pd.DataFrame(data)
    except Exception as e:
        # 시트가 없거나 연결 오류 시 빈 표 반환
        return pd.DataFrame()

# ==========================================
# 6단계: 전체 히스토리 내역 (분석 완료 시에만 노출 & 8대 항목 완벽 재현)
# ==========================================

# 🎯 분석 상태(analyzed)가 True일 때만 화면에 나타나도록 설정 (초기화 시 같이 사라짐)
if st.session_state.get('analyzed', False) and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📜 6단계: 전체 히스토리 내역")

    # [내부 함수] 구글 시트에서 '발주기록' 데이터를 읽어오는 전용 함수
    def load_v6_history_complete():
        try:
            sheet = get_sheet() 
            # 5단계에서 저장한 '발주기록' 시트를 불러옵니다.
            record_sheet = sheet.worksheet("발주기록")
            data = record_sheet.get_all_records()
            return pd.DataFrame(data)
        except Exception as e:
            # 시트가 없거나 연결 오류 시 빈 표 반환
            return pd.DataFrame()

    # 데이터 로딩 애니메이션
    with st.spinner('📡 구글 시트에서 히스토리 기록을 가져오는 중...'):
        df_hist = load_v6_history_complete()

    if not df_hist.empty:
        # 1. 상단 필터 UI (날짜 선택 및 검색)
        h_c1, h_c2 = st.columns([1, 2])
        
        # 한국 시간(KST) 기준 오늘 날짜 설정
        today_kst = datetime.now(KST).date()
        h_date = h_c1.date_input("🗓️ 조회 날짜 선택", today_kst, key="v6_final_date_input")
        h_search = h_c2.text_input("🔍 상품명 검색 (히스토리)", key="v6_final_search_input")

        # 2. 데이터 필터링 로직
        # '날짜' 컬럼에서 날짜 정보만 추출하여 달력과 비교
        if '날짜' in df_hist.columns:
            df_hist['날짜_only'] = pd.to_datetime(df_hist['날짜']).dt.date
            df_filtered = df_hist[df_hist['날짜_only'] == h_date].copy()
        else:
            # 혹시 컬럼명이 다를 경우를 대비한 방어 로직
            df_filtered = df_hist.copy()
        
        # 상품명 검색어 필터링
        if h_search:
            df_filtered = df_filtered[df_filtered['상품명'].astype(str).str.contains(h_search, case=False, na=False)]

        # 3. 🎯 사장님이 요청하신 8~9대 항목 순서 고정 표시
        # 저장 시점의 데이터: 날짜, 상태, 상품명, 옵션, 공급쳐상품명, 가용재고, 리오더수량, 추가발주수량, 권장발주량
        view_cols = ["날짜", "상태", "상품명", "옵션", "공급쳐상품명", "가용재고", "리오더수량", "추가발주수량", "권장발주량"]
        
        # 실제 시트에 존재하는 컬럼만 선별하여 에러 방지
        actual_show_cols = [c for c in view_cols if c in df_filtered.columns]

        # 4. 결과 출력
        if not df_filtered.empty:
            # 최신 기록이 위로 오도록 정렬하여 출력
            st.dataframe(
                df_filtered[actual_show_cols].sort_values(by="날짜", ascending=False), 
                use_container_width=True, 
                hide_index=True,
                column_config={
                    "가용재고": st.column_config.NumberColumn(format="%d"),
                    "리오더수량": st.column_config.NumberColumn(format="%d"),
                    "추가발주수량": st.column_config.NumberColumn(format="%d"),
                    "권장발주량": st.column_config.NumberColumn(format="%d")
                }
            )
            
            # 5. 📥 조회된 내역 CSV 다운로드 기능
            csv_data_v6 = df_filtered[actual_show_cols].to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button(
                label=f"📥 {h_date} 히스토리 다운로드 (엑셀용 CSV)",
                data=csv_data_v6,
                file_name=f"발주히스토리_{h_date}.csv",
                mime="text/csv",
                use_container_width=True
            )
        else:
            st.info(f"📅 {h_date} 날짜에는 저장된 기록이 없습니다. 날짜를 변경하거나 5단계에서 저장을 먼저 해주세요.")
    else:
        st.warning("📡 아직 저장된 히스토리 데이터가 없습니다. 5단계에서 '기록 저장' 버튼을 눌러주세요.")

# --- 6단계 끝 ---
    

# --- [🌙 탭 2: 동대문 사입 관리] ---
with tab2:
    st.subheader("🌙 동대문 사입 및 미납 관리")

    # 파일 업로드 (동대문 전용 키 사용)
    dong_file = st.file_uploader("동대문 주문 리스트 업로드", type=['xlsx', 'csv'], key="dong_tab_upload")

    if dong_file:
        # 1. 데이터 로드 및 초기화 (한 번만 실행)
        if "last_file_name" not in st.session_state or st.session_state.last_file_name != dong_file.name:
            with st.spinner('🚚 동대문 데이터를 분석 중입니다...'):
                df = pd.read_excel(dong_file) if not dong_file.name.endswith('.csv') else pd.read_csv(dong_file)
                df.columns = df.columns.str.strip()

                # 필수 컬럼 정의 및 없는 컬럼 자동 생성
                required_cols = ['선택', '품절', '상품명', '공급처', '공급처상품명', '정상재고', '가용재고', '판매수량', '발주수량', '가중율', '3일판매']
                for col in required_cols:
                    if col not in df.columns:
                        df[col] = 0 if col not in ['선택', '품절', '상품명', '공급처', '공급처상품명'] else ""
                
                # 데이터 타입 정제
                for col in ['정상재고', '가용재고', '3일판매']:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

                # 💡 [핵심 로직] 판매량 계산 및 가중율 자동 적용
                df['판매수량'] = (df['정상재고'] - df['가용재고']).clip(lower=0)
                # 판매량에 따른 가중치 (사장님 기존 로직 유지)
                df['가중율'] = df['판매수량'].apply(lambda n: 2.0 if n >= 10 else (1.5 if n >= 6 else (1.2 if n >= 3 else 1.0)))
                df['발주수량'] = (df['판매수량'] * df['가중율']).round(0).astype(int)
                
                # 선택(체크박스) 초기화
                df['선택'] = False

                st.session_state.df_dong_current = df[required_cols]
                st.session_state.last_file_name = dong_file.name

        # 2. 검색 및 편집 UI
        if "df_dong_current" in st.session_state:
            df_display = st.session_state.df_dong_current.copy()
            
            search_query = st.text_input("🔍 상품명 검색 (사입)", key="dong_search")
            if search_query:
                df_display = df_display[df_display['상품명'].astype(str).str.contains(search_query, case=False, na=False)]

            # 데이터 에디터 출력
            # '선택' 컬럼을 체크박스로 활용
            edited_df = st.data_editor(
                df_display, 
                use_container_width=True, 
                key="dong_editor",
                hide_index=True,
                column_config={
                    "선택": st.column_config.CheckboxColumn(help="발주 수량을 추가할 상품을 선택하세요"),
                    "가중율": st.column_config.NumberColumn(format="%.1f")
                }
            )

            st.divider()
            c1, c2, c3 = st.columns([1, 1, 1])
            
            # 수량 추가 로직
            add_val = c1.number_input("➕ 추가할 수량", value=1, min_value=1, key="dong_add_val")
            
            if c2.button("🚀 선택 상품 수량 더하기", use_container_width=True):
                # 에디터에서 '선택'이 True인 행의 인덱스 추출
                selected_indices = edited_df[edited_df['선택'] == True].index
                
                if not selected_indices.empty:
                    # 원본 세션 데이터에 수량 합산
                    for idx in selected_indices:
                        st.session_state.df_dong_current.at[idx, '발주수량'] += add_val
                    st.success(f"✅ {len(selected_indices)}개 항목에 {add_val}개씩 추가되었습니다.")
                    time.sleep(0.5)
                    st.rerun()
                else:
                    st.warning("선택된 상품이 없습니다.")

            # 3. 다운로드 버튼 (KST 시간 포함)
            file_time = datetime.now(KST).strftime('%m%d_%H%M')
            csv = edited_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            c3.download_button(
                label="📥 사입 리스트 다운로드",
                data=csv,
                file_name=f"동대문사입_{file_time}.csv",
                mime="text/csv",
                use_container_width=True
            )
    else:
        st.info("👆 동대문 주문 리스트(Excel/CSV)를 업로드해 주세요.")


# --- [🌙 탭 2: 동대문 사입 관리] ---
with tab2:
    st.subheader("🌙 동대문 사입 및 미납 관리")
    dong_file = st.file_uploader("동대문 주문 리스트 업로드", type=['xlsx', 'csv'], key="dong_tab_upload")
    
    # 1. 파일 업로드 및 데이터 처리
    if dong_file:
        if "last_file_name" not in st.session_state or st.session_state.last_file_name != dong_file.name:
            df = pd.read_excel(dong_file)
            df.columns = df.columns.str.strip()
            
            # [에러 방지] 엑셀에 컬럼이 없어도 강제로 생성
            required_cols = ['선택', '품절', '상품명', '공급처', '공급처상품명', '정상재고', '가용재고', '판매수량', '발주수량', '가중율', '3일판매']
            for col in required_cols:
                if col not in df.columns:
                    df[col] = 0 if col not in ['선택', '품절', '상품명', '공급처', '공급처상품명'] else ""
            
            # 수치형 변환
            for col in ['정상재고', '가용재고', '3일판매']:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
            # 계산 로직
            df['판매수량'] = (df['정상재고'] - df['가용재고']).clip(lower=0)
            df['가중율'] = df['판매수량'].apply(lambda n: 2.0 if n >= 10 else (1.5 if n >= 6 else (1.2 if n >= 3 else 1.0)))
            df['발주수량'] = (df['판매수량'] * df['가중율']).astype(int)
            
            st.session_state.df_dong_current = df[required_cols]
            st.session_state.last_file_name = dong_file.name

        # 2. 화면 출력
        df_display = st.session_state.df_dong_current.copy()
        
        # [검색창]
        c1, c2 = st.columns([1, 2])
        search_query = c2.text_input("상품명 검색")
        if search_query:
            df_display = df_display[df_display['상품명'].astype(str).str.contains(search_query, case=False, na=False)]

        # [데이터 편집기]
        df_display['선택'] = df_display['선택'].astype(bool)
        edited_df = st.data_editor(
            df_display, use_container_width=True, key="final_editor",
            column_config={"선택": st.column_config.CheckboxColumn("선택", width="small")}
        )

        # [버튼]
        st.divider()
        col1, col2, col3 = st.columns(3)
        add_val = col1.number_input("추가 수량", value=1, min_value=1)
        
        if col2.button("🚀 선택한 상품 수량 더하기"):
            # 선택된 인덱스만 발주수량 업데이트
            selected = edited_df[edited_df['선택'] == True].index
            for idx in selected:
                st.session_state.df_dong_current.at[idx, '발주수량'] += add_val
            st.rerun()
            
        csv = edited_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
        col3.download_button("📥 엑셀 다운로드", csv, "사입리스트.csv", "text/csv")
