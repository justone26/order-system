import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- [1. 구글 시트 연동 함수] ---
def get_sheet():
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        spreadsheet_key = "1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg"
        return client.open_by_key(spreadsheet_key)
    except Exception as e:
        st.error(f"구글 시트 연결 실패: {e}")
        return None

def get_recent_logs():
    """입고 날짜/수량 체크를 위해 history 시트에서 최근 3건 요약 로드"""
    try:
        sh = get_sheet()
        hist_sheet = sh.worksheet("history")
        data = hist_sheet.get_all_records()
        if not data: return {}
        df_h = pd.DataFrame(data)
        # 리오더입고수량이 있는 행만 필터링
        df_in = df_h[pd.to_numeric(df_h['리오더입고수량'], errors='coerce') > 0].copy()
        log_dict = {}
        for (name, opt), group in df_in.groupby(['상품명', '옵션']):
            recent = group.sort_values(by='저장시간', ascending=False).head(3)
            logs = [f"{str(r['저장시간'])[5:10]}({int(r['리오더입고수량'])}개)" for _, r in recent.iterrows()]
            log_dict[(str(name).strip(), str(opt).strip())] = " / ".join(logs)
        return log_dict
    except: return {}

# --- [2. 앱 설정 및 탭 정의] ---
st.set_page_config(layout="wide", page_title="재고 관리 시스템")
st.title("📦 재고 관리 및 발주 시스템")

# 탭을 최상단에 정의하여 NameError 방지
tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 관리"])

if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'in_logs' not in st.session_state: st.session_state.in_logs = {}

def get_idx(cols, keywords):
    for key in keywords:
        for i, c in enumerate(cols):
            if key in str(c): return i
    return 0

# --- [🏭 탭 1: 제작 상품 관리] ---
with tab1:
    # 1. 업로드 및 리오더 수량 복구 (초기화 방지)
    uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'xls', 'csv'])
    if uploaded_file is not None and st.session_state.df_raw is None:
        df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
        df.columns = df.columns.str.strip()
        df = df.loc[:, ~df.columns.duplicated()]

        # [수정 1] 구글 시트에서 기존 리오더 수량 불러오기 (초기화 방지)
        try:
            sh = get_sheet()
            gs_data = pd.DataFrame(sh.sheet1.get_all_records())
            if not gs_data.empty and '리오더 수량' in gs_data.columns:
                # 상품명+옵션 조합으로 매칭
                df['match_key'] = df.iloc[:, get_idx(df.columns, ['상품명'])].astype(str).str.strip() + \
                                  df.iloc[:, get_idx(df.columns, ['옵션'])].astype(str).str.strip()
                gs_data['match_key'] = gs_data['상품명'].astype(str).str.strip() + gs_data['옵션'].astype(str).str.strip()
                reorder_map = gs_data.set_index('match_key')['리오더 수량'].to_dict()
                df['입고예정수량(리오더)'] = df['match_key'].map(reorder_map).fillna(0).astype(int)
                df.drop(columns=['match_key'], inplace=True)
            else:
                df['입고예정수량(리오더)'] = 0
            st.session_state.in_logs = get_recent_logs()
        except:
            df['입고예정수량(리오더)'] = 0

        st.session_state.df_raw = df
        st.rerun()

    if st.session_state.df_raw is not None:
        cols = st.session_state.df_raw.columns.tolist()

        # [1단계: 매핑]
        st.subheader("⚙️ 1단계: 자동 매핑 설정")
        c1, c2 = st.columns(2)
        sold_out = c1.selectbox("품절 여부", cols, index=get_idx(cols, ['품절', '판매중단']))
        vendor = c1.selectbox("공급처", cols, index=get_idx(cols, ['공급처', '업체명']))
        item = c1.selectbox("상품명", cols, index=get_idx(cols, ['상품명', '상품']))
        option = c1.selectbox("옵션", cols, index=get_idx(cols, ['옵션']))
        vendor_opt = c1.selectbox("공급처옵션", cols, index=get_idx(cols, ['공급처옵션', '거래처옵션']))
        stock = c2.selectbox("정상재고", cols, index=get_idx(cols, ['정상재고', '재고']))
        avail = c2.selectbox("가용재고", cols, index=get_idx(cols, ['가용재고', '가용']))
        t3day = c2.selectbox("3일 발주 합계", cols, index=get_idx(cols, ['3일', '최근3일']))
        t1week = c2.selectbox("1주 발주 합계", cols, index=get_idx(cols, ['1주', '7일', '최근7일']))

        # [2단계: 파라미터]
        st.subheader("⚙️ 2단계: 파라미터 설정")
        l1, l2 = st.columns(2)
        lead_time = l1.number_input("리드타임", value=10)
        safety_stock = l2.number_input("안전재고", value=7)

        # [3단계: 분석]
        st.subheader("⚙️ 3단계: 분석 실행")
        if st.button("🚀 분석 실행"):
            df = st.session_state.df_raw
            df['일일 판매량'] = (pd.to_numeric(df[t3day], errors='coerce').fillna(0) / 3).round(1)
            df['권장 발주량'] = ((df['일일 판매량'] * (lead_time + safety_stock)) - 
                             (pd.to_numeric(df[avail], errors='coerce').fillna(0) + df['입고예정수량(리오더)'])).clip(lower=0).round(0).astype(int)
            
            # [수정 2] 품절건은 리오더 발주에서 제외 (권장 발주량 0)
            df.loc[df[sold_out].astype(str).str.contains('품절', na=False), '권장 발주량'] = 0
            st.rerun()

        # [4단계: 실시간 입고 관리 및 날짜 체크 셀]
        st.subheader("📊 4단계: 검색 및 실시간 입고 관리")
        
        # [수정 3] 입고 날짜 체크용 데이터 매칭
        def match_log(row):
            key = (str(row[item]).strip(), str(row[option]).strip())
            return st.session_state.in_logs.get(key, "-")
        
        st.session_state.df_raw['최근입고기록(날짜/수량)'] = st.session_state.df_raw.apply(match_log, axis=1)

        f1, f2 = st.columns([3, 1])
        search = f1.text_input("🔍 상품명 검색")
        filter_mode = f2.selectbox("품절 필터", ["전체보기", "품절만", "정상만"], index=2)
        
        df_disp = st.session_state.df_raw.copy()
        if filter_mode == "품절만": df_disp = df_disp[df_disp[sold_out].astype(str).str.contains('품절', na=False)]
        elif filter_mode == "정상만": df_disp = df_disp[~df_disp[sold_out].astype(str).str.contains('품절', na=False)]
        if search: df_disp = df_disp[df_disp[item].astype(str).str.contains(search, na=False)]

        if "리오더입고수량" not in df_disp.columns: df_disp["리오더입고수량"] = 0
        edit_cols = [sold_out, vendor, item, option, vendor_opt, stock, avail, "입고예정수량(리오더)", "리오더입고수량", "최근입고기록(날짜/수량)", '권장 발주량']
        
        # 실시간 입고 처리 핸들러
        def on_edit_prod():
            changes = st.session_state["main_editor"]["edited_rows"]
            for idx_str, change in changes.items():
                idx = int(idx_str)
                orig_idx = df_disp.index[idx]
                if "리오더입고수량" in change:
                    in_qty = int(change["리오더입고수량"])
                    st.session_state.df_raw.at[orig_idx, "입고예정수량(리오더)"] -= in_qty
                    # history 시트에 기록 저장
                    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    sh = get_sheet()
                    sh.worksheet("history").append_row([now_str, str(df_disp.at[orig_idx, item]), str(df_disp.at[orig_idx, option]), in_qty])
            
            # 구글 시트 Sheet1에 현재 리오더 수량 실시간 저장 (초기화 방지)
            save_df = st.session_state.df_raw[[item, option, "입고예정수량(리오더)"]].copy()
            save_df.columns = ['상품명', '옵션', '리오더 수량']
            sh = get_sheet()
            sh.sheet1.update('A1', [save_df.columns.values.tolist()] + save_df.values.tolist())
            st.session_state.in_logs = get_recent_logs()
            st.rerun()

        st.data_editor(df_disp[edit_cols], use_container_width=True, key="main_editor", on_change=on_edit_prod)

        # [5단계: 발주 요약]
        st.subheader("📋 5단계: 최종 발주 리스트 (품절 제외)")
        if '권장 발주량' in st.session_state.df_raw.columns:
            to_order = st.session_state.df_raw[st.session_state.df_raw['권장 발주량'] > 0]
            st.dataframe(to_order[edit_cols], use_container_width=True)
            
            c_save, c_down = st.columns(2)
            if c_save.button("💾 발주 내역 시트 저장"):
                now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                sh = get_sheet()
                rows = [[now_str] + row for row in to_order[[item, option, '권장 발주량']].values.tolist()]
                sh.worksheet("history").append_rows(rows)
                st.success("발주 내역이 히스토리에 저장되었습니다.")
            
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                to_order.to_excel(writer, index=False)
            c_down.download_button("📥 최종 데이터 다운로드", data=output.getvalue(), file_name="최종발주서.xlsx")

# --- [🌙 탭 2: 동대문 사입 관리] ---
with tab2:
    st.subheader("🌙 동대문 사입 및 미납 관리")
    dong_file = st.file_uploader("동대문 주문 리스트 업로드", type=['xlsx', 'csv'], key="dong_tab_upload")
    if dong_file:
        if "last_file_name" not in st.session_state or st.session_state.last_file_name != dong_file.name:
            df = pd.read_excel(dong_file) if not dong_file.name.endswith('.csv') else pd.read_csv(dong_file)
            df.columns = df.columns.str.strip()
            required_cols = ['선택', '품절', '상품명', '공급처', '공급처상품명', '정상재고', '가용재고', '판매수량', '발주수량', '가중율', '3일판매']
            for col in required_cols:
                if col not in df.columns: df[col] = 0 if col not in ['선택', '품절', '상품명', '공급처', '공급처상품명'] else ""
            for col in ['정상재고', '가용재고']: df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            df['판매수량'] = (df['정상재고'] - df['가용재고']).clip(lower=0)
            df['가중율'] = df['판매수량'].apply(lambda n: 2.0 if n >= 10 else (1.5 if n >= 6 else (1.2 if n >= 3 else 1.0)))
            df['발주수량'] = (df['판매수량'] * df['가중율']).astype(int)
            st.session_state.df_dong_current = df[required_cols]
            st.session_state.last_file_name = dong_file.name

        df_display = st.session_state.df_dong_current.copy()
        search_query = st.text_input("상품명 검색 (사입)")
        if search_query: df_display = df_display[df_display['상품명'].astype(str).str.contains(search_query, case=False, na=False)]
        
        df_display['선택'] = df_display['선택'].astype(bool)
        edited_dong = st.data_editor(df_display, use_container_width=True, key="dong_editor")
        
        st.divider()
        c1, c2, c3 = st.columns(3)
        add_val = c1.number_input("추가 수량", value=1, min_value=1)
        if c2.button("🚀 선택 상품 수량 더하기"):
            # editor에서 선택된 인덱스 가져오기
            selected_indices = edited_dong[edited_dong['선택'] == True].index
            for idx in selected_indices:
                st.session_state.df_dong_current.at[idx, '발주수량'] += add_val
            st.rerun()
        
        csv = edited_dong.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
        c3.download_button("📥 엑셀 다운로드", csv, "사입리스트.csv", "text/csv")

    st.subheader("📜 6단계: 과거 기록 조회")
    if st.button("🔄 기록 새로고침"):
        sh = get_sheet()
        hist_df = pd.DataFrame(sh.worksheet("history").get_all_records())
        st.dataframe(hist_df.sort_values(by='저장시간', ascending=False), use_container_width=True)
