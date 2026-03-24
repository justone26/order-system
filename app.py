import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- [1. 기본 함수 설정] ---
def get_sheet():
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        spreadsheet_key = "1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg"
        return client.open_by_key(spreadsheet_key)
    except: return None

def load_history_from_gsheet():
    try:
        spreadsheet = get_sheet()
        hist_sheet = spreadsheet.worksheet("history")
        data = hist_sheet.get_all_records()
        return pd.DataFrame(data)
    except: return pd.DataFrame()

def make_match_key(name, opt):
    return str(name).strip().replace(" ", "").upper() + str(opt).strip().replace(" ", "").upper()

def save_reorder_data(new_work_df):
    try:
        spreadsheet = get_sheet()
        if not spreadsheet: return False
        
        sheet = spreadsheet.sheet1
        
        # 1. 구글 시트에 이미 있는 데이터 전체 읽어오기
        raw_gs_data = sheet.get_all_records()
        if raw_gs_data:
            gs_df = pd.DataFrame(raw_gs_data)
        else:
            # 시트가 비어있을 경우 기본 틀 생성
            gs_df = pd.DataFrame(columns=['상품명', '옵션', '리오더 수량'])

        # 2. 비교를 위한 '매칭 키' 생성 함수
        def make_key(df_in):
            # 상품명과 옵션을 합쳐서 고유한 열쇠를 만듭니다 (공백/대소문자 무시)
            return df_in['상품명'].astype(str).str.strip().str.replace(" ", "").str.upper() + \
                   df_in['옵션'].astype(str).str.strip().str.replace(" ", "").str.upper()

        # 기존 데이터에 키 추가
        if not gs_df.empty:
            gs_df['match_key'] = make_key(gs_df)
        else:
            gs_df['match_key'] = ""

        # 새로 들어온 데이터(엑셀 등)에 키 추가
        new_work_df['match_key'] = make_key(new_work_df)
        
        if not gs_df.empty:
            gs_df['리오더 수량'] = pd.to_numeric(gs_df['리오더 수량'], errors='coerce').fillna(0)
        
        # 3. 데이터 병합 (Upsert 로직)
        for _, row in new_work_df.iterrows():
            target_key = row['match_key']
            
            # 이미 시트에 있는 상품이면? -> '리오더 수량'만 업데이트
            if target_key in gs_df['match_key'].values:
                gs_df.loc[gs_df['match_key'] == target_key, '리오더 수량'] += row['리오더 수량']
            # 처음 보는 상품(다른 업체 등)이면? -> 아래에 새로 추가
            else:
                # 필요한 컬럼만 추출해서 합치기
                new_item = pd.DataFrame([{
                    '상품명': row['상품명'],
                    '옵션': row['옵션'],
                    '리오더 수량': row['리오더 수량']
                }])
                gs_df = pd.concat([gs_df, new_item], ignore_index=True)

        # 4. 불필요한 키 삭제 및 정리
        final_df = gs_df.drop(columns=['match_key'], errors='ignore').fillna(0)
        
        # 중복된 행이 혹시 생기면 마지막 것만 남기기 (안전장치)
        final_df = final_df.drop_duplicates(subset=['상품명', '옵션'], keep='last')

        # 5. 시트 최종 업데이트
        sheet.clear()
        # 헤더 포함 전체 데이터 쓰기
        sheet.update([final_df.columns.values.tolist()] + final_df.values.tolist())
        return True
    except Exception as e:
        st.error(f"⚠️ 데이터 누적 저장 중 오류 발생: {e}")
        return False

def save_history_to_gsheet(df, log_type="입고"):
    try:
        spreadsheet = get_sheet()
        if not spreadsheet: return False
        
        # history 시트 가져오기 또는 생성
        try:
            hist_sheet = spreadsheet.worksheet("history")
        except:
            hist_sheet = spreadsheet.add_worksheet(title="history", rows="1000", cols="20")
            hist_sheet.append_row(["저장시간", "구분", "상품명", "옵션", "수량"])
        
        # 현재 시간 생성
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # 저장할 데이터 구성 (사장님 시트 양식: 저장시간, 구분, 상품명, 옵션, 수량)
        # df에는 보통 [상품명, 옵션, 수량]만 넘어오므로 앞에 시간과 구분을 붙여줍니다.
        rows_to_add = []
        for row in df.values.tolist():
            rows_to_add.append([now_str, log_type] + [str(x) for x in row])
        
        if rows_to_add:
            hist_sheet.append_rows(rows_to_add)
            return True
        return False
    except Exception as e:
        # 에러 발생 시 화면에 표시 (사장님 확인용)
        st.error(f"히스토리 저장 실패: {e}")
        return False

# 아래 두 개는 사장님 기존 코드 그대로 쓰셔도 완벽합니다!
def find_idx(cols, target_keywords):
    for keyword in target_keywords:
        for i, col in enumerate(cols):
            if keyword in str(col): return i
    return 0

def safe_num(val):
    res = pd.to_numeric(val, errors='coerce')
    if isinstance(res, pd.Series): return res.fillna(0)
    return 0 if pd.isna(res) else res

# --- [2. 앱 초기 설정] ---
st.set_page_config(layout="wide", page_title="저스트원 재고관리")
st.title("🏭 저스트원 통합 재고 관리 시스템")

if "extra_order_dict" not in st.session_state: st.session_state.extra_order_dict = {}
if 'analyzed' not in st.session_state: st.session_state.analyzed = False

tab1, tab2 = st.tabs(["✂️ 제작 상품 관리", "🌙 동대문 상품 관리"])

# --- [탭 1: 제작 상품 관리] ---
with tab1:
    uploaded_file = st.file_uploader("엑셀 파일을 올려주세요", type=['xlsx', 'xls', 'csv'], key="t1_up")
    
    if st.button("📂 구글 시트 데이터 로드", use_container_width=True):
        spreadsheet = get_sheet()
        if spreadsheet:
            try:
                # 1. 첫 번째 워크시트 선택
                sheet = spreadsheet.get_worksheet(0)
                
                # 2. 🔥 중요: get_all_records() 대신 get_all_values() 사용 후 직접 처리
                # 이 방식이 헤더와 데이터를 가장 확실하게 구분합니다.
                raw_data = sheet.get_all_values()
                
                if len(raw_data) > 1:
                    # 첫 줄을 컬럼명으로, 나머지를 데이터로 분리
                    header = [str(h).strip() for h in raw_data[0]] # 공백 제거
                    content = raw_data[1:]
                    
                    df_tmp = pd.DataFrame(content, columns=header)
                    
                    # 3. 열 이름 중복 방지 (Streamlit 에러 방어)
                    new_cols = []
                    for i, col in enumerate(df_tmp.columns):
                        if not col or col in new_cols:
                            new_cols.append(f"열_{i}") # 이름 없거나 중복이면 강제 부여
                        else:
                            new_cols.append(col)
                    df_tmp.columns = new_cols
                    
                    # 4. 세션에 저장
                    st.session_state.df_raw = df_tmp.copy()
                    st.session_state.analyzed = False
                    st.success(f"✅ {len(df_tmp)}개 항목의 열 이름을 성공적으로 가져왔습니다!")
                    st.rerun()
                else:
                    st.warning("⚠️ 시트에 데이터가 부족합니다. 최소한 제목줄과 데이터 한 줄은 있어야 합니다.")
            except Exception as e:
                st.error(f"❌ 데이터 로드 중 오류: {e}")
        else:
            st.error("❌ 시트 연결 실패! 공유 권한을 확인해 주세요.")
            
    if uploaded_file:
        if 'df_raw' not in st.session_state or st.session_state.get('last_fn') != uploaded_file.name:
            df_new = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
            df_new.columns = df_new.columns.str.strip()
            # 리오더 수량 매핑
            try:
                sheet = get_sheet().sheet1
                gs_df = pd.DataFrame(sheet.get_all_records())
                if not gs_df.empty and '리오더 수량' in gs_df.columns:
                    t_item = next((c for c in df_new.columns if '상품명' in c), df_new.columns[0])
                    t_opt = next((c for c in df_new.columns if '옵션' in c), df_new.columns[1])
                    df_new['k_tmp'] = df_new.apply(lambda r: make_match_key(r[t_item], r[t_opt]), axis=1)
                    gs_df['k_tmp'] = gs_df.apply(lambda r: make_match_key(r['상품명'], r['옵션']), axis=1)
                    rmap = gs_df.set_index('k_tmp')['리오더 수량'].to_dict()
                    df_new['리오더 수량'] = df_new['k_tmp'].map(rmap).fillna(0).astype(int)
                    df_new.drop(columns=['k_tmp'], inplace=True)
                else: df_new['리오더 수량'] = 0
            except: df_new['리오더 수량'] = 0
            st.session_state.df_raw = df_new
            st.session_state.last_fn = uploaded_file.name
            st.session_state.analyzed = False
            st.rerun()

    if st.session_state.get('df_raw') is not None:
        df_curr = st.session_state.df_raw
        cols = df_curr.columns.tolist()
        
        st.subheader("⚙️ 3단계: 매핑 설정")
        c_l, c_r = st.columns(2)
        with c_l:
            sold_out = st.selectbox("품절 여부", cols, index=find_idx(cols, ['품절']))
            vendor = st.selectbox("공급처", cols, index=find_idx(cols, ['공급처']))
            v_item = st.selectbox("공급처 상품명", cols, index=find_idx(cols, ['공급처상품명']))
            item = st.selectbox("상품명", cols, index=find_idx(cols, ['상품명']))
            option = st.selectbox("옵션", cols, index=find_idx(cols, ['옵션']))
        with c_r:
            reg_date = st.selectbox("등록일", cols, index=find_idx(cols, ['등록일']))
            stock = st.selectbox("정상재고", cols, index=find_idx(cols, ['정상재고']))
            avail = st.selectbox("가용재고", cols, index=find_idx(cols, ['가용재고']))
            t3day = st.selectbox("3일 발주합계", cols, index=find_idx(cols, ['3일']))
            t7day = st.selectbox("7일 발주합계", cols, index=find_idx(cols, ['7일', '1주']))

        lt = st.number_input("리드타임 (일)", value=7)
        ss = st.number_input("안전재고 (일 수)", value=3)
        if st.button("📊 분석 실행", use_container_width=True):
            st.session_state.analyzed = True
            st.rerun()
            
# --- [4단계: 데이터 편집 및 재고 관리] ---
        st.divider()
        st.subheader("📊 4단계: 데이터 편집 및 재고 관리")
        
        # 1. 원본 데이터 복사 및 입고수량 칸 생성 (값이 유지되도록 세션에 저장)
        df_work = st.session_state.df_raw.copy()
        if "리오더입고수량" not in st.session_state.df_raw.columns:
            st.session_state.df_raw["리오더입고수량"] = 0

        f_c1, f_c2, f_c3 = st.columns([2, 1, 1])
        search_q = f_c1.text_input("🔍 상품명 검색", key="search_v4")
        filter_m = f_c2.selectbox("품절 필터", ["전체보기", "정상만", "품절만"], index=1, key="filter_v4")
        hist_date_4 = f_c3.date_input("🗓️ 입고 기록 확인 날짜", datetime.now(), key="date_v4")

        def simple_key(n): return str(n).strip().replace(" ", "").upper() if not pd.isna(n) else ""
        df_work['unique_key'] = df_work[item].apply(simple_key) + df_work[option].apply(simple_key)

        past_hist = load_history_from_gsheet()
        df_work['과거 리오더입고'] = 0
        
        # 💡 [핵심] 입고 수량 수치가 화면에 유지되도록 세션 데이터를 연결
        df_work['리오더입고수량'] = st.session_state.df_raw['리오더입고수량']
        
        if not past_hist.empty and '저장시간' in past_hist.columns and '구분' in past_hist.columns:
            try:
                past_hist['날짜_only'] = pd.to_datetime(past_hist['저장시간']).dt.date
                t_hist = past_hist[(past_hist['날짜_only'] == hist_date_4) & (past_hist['구분'] == "입고")].copy()
                if not t_hist.empty:
                    t_hist['k_tmp'] = t_hist['상품명'].apply(simple_key) + t_hist['옵션'].apply(simple_key)
                    in_map = t_hist.groupby('k_tmp')['수량'].sum().to_dict()
                    df_work['과거 리오더입고'] = df_work['unique_key'].map(in_map).fillna(0).astype(int)
            except:
                pass

        v7 = safe_num(df_work[t7day]); v3 = safe_num(df_work[t3day])
        df_work['일판매량'] = (v7 / 7 if v7.sum() > 0 else v3 / 3).round(0).astype(int)
        df_work['권장발주량'] = ((df_work['일판매량'] * (lt + ss)) - (safe_num(df_work[avail]) + safe_num(df_work['리오더 수량']))).clip(lower=0).astype(int)

        if filter_m == "정상만": df_work = df_work[~df_work[sold_out].astype(str).str.contains('품절', na=False)]
        elif filter_m == "품절만": df_work = df_work[df_work[sold_out].astype(str).str.contains('품절', na=False)]
        if search_q: df_work = df_work[df_work[item].astype(str).str.contains(search_q, case=False, na=False)]

        valid_cols = [sold_out, vendor, v_item, item, option, stock, avail, "리오더 수량", "리오더입고수량", "과거 리오더입고", t3day, "일판매량", "권장발주량"]
        
        # 💡 [핵심] 실시간 저장 및 수치 유지 로직 (st.rerun 제거 버전)
        def on_edit_4():
            changes = st.session_state["editor_v4"]["edited_rows"]
            for r_idx_str, change in changes.items():
                orig_idx = df_work.index[int(r_idx_str)]
                
                if "리오더 수량" in change:
                    st.session_state.df_raw.at[orig_idx, "리오더 수량"] = int(change["리오더 수량"])
                
                if "리오더입고수량" in change:
                    new_in_qty = int(change["리오더입고수량"])
                    old_in_qty = st.session_state.df_raw.at[orig_idx, "리오더입고수량"]
                    
                    if new_in_qty > old_in_qty:
                        diff = new_in_qty - old_in_qty
                        curr_reorder = int(st.session_state.df_raw.at[orig_idx, "리오더 수량"])
                        st.session_state.df_raw.at[orig_idx, "리오더 수량"] = max(0, curr_reorder - diff)
                        save_history_to_gsheet(pd.DataFrame([[df_work.at[orig_idx, item], df_work.at[orig_idx, option], diff]], columns=['상품명', '옵션', '수량']), log_type="입고")
                    
                    st.session_state.df_raw.at[orig_idx, "리오더입고수량"] = new_in_qty

            save_reorder_data(st.session_state.df_raw[[item, option, '리오더 수량']].rename(columns={item:'상품명', option:'옵션'}))
            # ✅ st.rerun()을 삭제했습니다. 스트림릿이 알아서 새로고침합니다.

        st.data_editor(df_work[valid_cols], use_container_width=True, key="editor_v4", on_change=on_edit_4, hide_index=True)

# --- [5단계: 최종 발주 리스트 요약 - 수치 유지 및 엑셀 연동 버전] ---
        st.divider()
        st.subheader("📋 5단계: 최종 발주 리스트 요약")
        
        # 1. 4단계 데이터 가져오기
        df_final_sync = st.session_state.df_raw.copy()
        
        # 💡 [핵심] 추가발주수량이 세션에 없으면 새로 만듭니다. (값 유지용)
        if 'add_order_dict' not in st.session_state:
            st.session_state.add_order_dict = {} # {인덱스: 수량} 형태로 저장

        c5_1, c5_2 = st.columns([2, 1])
        s_filter = c5_1.selectbox("🎯 상태 필터", ["🚨긴급 + ⚠️주의 우선", "🚨 긴급만 보기", "✅ 정상 포함 전체보기"], index=0, key="s_filter_v5")
        hist_date_5 = c5_2.date_input("🗓️ 입고 기록 확인 날짜 (연동)", value=hist_date_4, key="date_v5")

        to_order = df_final_sync.copy()
        
        # 💡 [핵심] 세션에 저장된 추가발주수량을 데이터프레임에 매핑합니다.
        to_order['unique_idx'] = to_order.index # 고유 인덱스 활용
        to_order['추가발주수량'] = to_order['unique_idx'].map(st.session_state.add_order_dict).fillna(0).astype(int)

        # 재계산 로직 (기존과 동일)
        v7_f = safe_num(to_order[t7day]); v3_f = safe_num(to_order[t3day])
        to_order['일판매량'] = (v7_f / 7 if v7_f.sum() > 0 else v3_f / 3).round(0).astype(int)
        to_order['권장발주량'] = ((to_order['일판매량'] * (lt + ss)) - (safe_num(to_order[avail]) + safe_num(to_order['리오더 수량']))).clip(lower=0).astype(int)
        
        def get_final_status(r):
            total_stock = safe_num(r[avail]) + safe_num(r['리오더 수량'])
            daily = r['일판매량']
            if daily > 0:
                if total_stock < (daily * 3): return "🚨 긴급"
                if total_stock < (daily * 5): return "⚠️ 주의"
            return "✅ 정상"
        
        to_order['상태'] = to_order.apply(get_final_status, axis=1)
        to_order = to_order.sort_values(by='상태')

        if s_filter == "🚨긴급 + ⚠️주의 우선": 
            to_order = to_order[to_order['상태'].isin(["🚨 긴급", "⚠️ 주의"]) | (to_order['권장발주량'] > 0)]
        elif s_filter == "🚨 긴급만 보기": 
            to_order = to_order[to_order['상태'] == "🚨 긴급"]

        disp_final = ["상태", item, option, vendor, avail, "리오더 수량", "추가발주수량", "권장발주량"]
        actual_cols = [c for c in disp_final if c in to_order.columns]
        
        # 💡 [핵심] 5단계 편집 시 값을 세션에 저장하는 콜백
        def on_edit_5():
            edits = st.session_state["editor_v5"]["edited_rows"]
            for r_idx_str, change in edits.items():
                orig_idx = to_order.index[int(r_idx_str)]
                if "추가발주수량" in change:
                    new_val = int(change["추가발주수량"])
                    # 1. 세션 딕셔너리에 저장 (화면 유지용)
                    st.session_state.add_order_dict[orig_idx] = new_val
                    
                    # 2. 실제 리오더 수량에도 반영 (선택사항: 원본 데이터를 바꾸고 싶을 때만)
                    # st.session_state.df_raw.at[orig_idx, "리오더 수량"] += (new_val - 기존값) 로직이 필요할 수 있음

        st.data_editor(to_order[actual_cols], use_container_width=True, key="editor_v5", on_change=on_edit_5, hide_index=True)

        b1, b2 = st.columns(2)
        if b1.button("💾 구글 시트에 최종 발주 기록 저장", use_container_width=True, type="primary"):
            # 권장발주량 + 추가발주수량이 있는 것들 저장
            to_order['최종발주량'] = to_order['권장발주량'] + to_order['추가발주수량']
            order_final = to_order[to_order['최종발주량'] > 0].copy()
            if not order_final.empty:
                if save_history_to_gsheet(order_final[[item, option, '최종발주량']], log_type="발주"):
                    st.success("✅ 발주 내역 저장 완료!")
                    # 저장 후에는 추가발주수량 초기화 여부 결정 (보통은 비웁니다)
                    # st.session_state.add_order_dict = {} 
                    st.rerun()
            else: st.warning("발주할 항목이 없습니다.")

        # 💡 [개선] 이제 다운로드 받는 CSV에 사장님이 입력한 '추가발주수량'이 포함됩니다!
        csv_v5 = to_order[actual_cols].to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
        b2.download_button("📥 엑셀(CSV) 다운로드", data=csv_v5, file_name=f"발주서_{datetime.now().strftime('%m%d')}.csv", use_container_width=True)

# --- [6단계: 기록 통합 조회 - KeyError 방어 강화] ---
        st.divider()
        st.subheader("📜 6단계: 제작 상품 입고 및 발주 히스토리")
        past_hist = load_history_from_gsheet() 
    
        c6_1, c6_2 = st.columns(2)
        start_d = c6_1.date_input("📅 조회 시작일", datetime.now() - timedelta(days=7), key="s_date_v6")
        end_d = c6_2.date_input("📅 조회 종료일", datetime.now(), key="e_date_v6")

        if not past_hist.empty:
            # 1. 날짜 변환 및 필터링
            if '저장시간' in past_hist.columns:
                past_hist['날짜_dt'] = pd.to_datetime(past_hist['저장시간'], errors='coerce').dt.date
                df_h = past_hist[(past_hist['날짜_dt'] >= start_d) & (past_hist['날짜_dt'] <= end_d)].copy()
                
                # 2. '구분' 열 존재 여부 체크 (에러 방지 핵심)
                if '구분' not in df_h.columns:
                    st.info("💡 히스토리 시트에 '구분' 항목이 없습니다. 데이터를 먼저 저장해 주세요.")
                else:
                    t_in, t_out = st.tabs(["📥 입고 완료 내역", "📤 발주 진행 내역"])
                    
                    with t_in:
                        in_df = df_h[df_h['구분'] == "입고"]
                        if not in_df.empty:
                            st.dataframe(in_df[['저장시간', '상품명', '옵션', '수량']], use_container_width=True, hide_index=True)
                            # 합계 계산 시 컬럼 체크
                            sum_cols = [c for c in ['상품명', '옵션', '수량'] if c in in_df.columns]
                            st.markdown("##### 📊 품목별 입고 합계")
                            st.table(in_df.groupby(['상품명', '옵션'])['수량'].sum().reset_index())
                        else:
                            st.write("선택한 기간에 '입고' 내역이 없습니다.")
                            
                    with t_out:
                        out_df = df_h[df_h['구분'] == "발주"]
                        if not out_df.empty:
                            st.dataframe(out_df[['저장시간', '상품명', '옵션', '수량']], use_container_width=True, hide_index=True)
                            st.markdown("##### 📊 품목별 발주 합계")
                            st.table(out_df.groupby(['상품명', '옵션'])['수량'].sum().reset_index())
                        else:
                            st.write("선택한 기간에 '발주' 내역이 없습니다.")
            else:
                st.warning("⚠️ 히스토리 시트에 '저장시간' 열이 없습니다.")
        else:
            st.info("💡 아직 구글 시트에 누적된 기록이 없습니다. 먼저 '발주 기록 저장'이나 '입고'를 진행해 주세요.")

# --- [🌙 탭 2: 동대문 사입 관리] ---
with tab2:
    st.subheader("🌙 동대문 사입 및 미납 관리")

    dong_file = st.file_uploader("동대문 주문 리스트 업로드", type=['xlsx', 'xls', 'csv'], key="dong_tab_upload")

    if dong_file:
        # 파일이 새로 올라왔을 때만 데이터 처리
        if "last_file_name" not in st.session_state or st.session_state.last_file_name != dong_file.name:
            # 엑셀/CSV 구분해서 읽기
            if dong_file.name.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(dong_file)
            else:
                df = pd.read_csv(dong_file)
            
            df.columns = df.columns.str.strip()
            
            # 필수 컬럼 체크 및 생성
            required_cols = ['선택', '품절', '상품명', '공급처', '공급처상품명', '정상재고', '가용재고', '판매수량', '발주수량', '가중율', '3일판매']
            for col in required_cols:
                if col not in df.columns:
                    if col in ['선택', '품절', '상품명', '공급처', '공급처상품명']:
                        df[col] = ""
                    else:
                        df[col] = 0
            
            # 숫자형 변환
            for col in ['정상재고', '가용재고', '3일판매']:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
            # 동대문 전용 발주 로직 적용
            df['판매수량'] = (df['정상재고'] - df['가용재고']).clip(lower=0)
            # 판매량에 따른 가중율 (10개 이상 2배, 6개 이상 1.5배 등)
            df['가중율'] = df['판매수량'].apply(lambda n: 2.0 if n >= 10 else (1.5 if n >= 6 else (1.2 if n >= 3 else 1.0)))
            df['발주수량'] = (df['판매수량'] * df['가중율']).astype(int)
            
            st.session_state.df_dong_current = df[required_cols]
            st.session_state.last_file_name = dong_file.name

        # 화면 출력 부분
        if "df_dong_current" in st.session_state:
            df_display = st.session_state.df_dong_current.copy()
            
            # 검색 기능
            search_query = st.text_input("🔍 상품명 검색 (사입)")
            if search_query:
                df_display = df_display[df_display['상품명'].astype(str).str.contains(search_query, case=False, na=False)]
            
            # 데이터 에디터 (선택 및 수량 수정 가능)
            df_display['선택'] = df_display['선택'].apply(lambda x: True if x is True or x == "True" else False)
            
            edited_df = st.data_editor(
                df_display, 
                use_container_width=True, 
                key="dong_editor",
                column_config={
                    "선택": st.column_config.CheckboxColumn("선택", default=False),
                    "발주수량": st.column_config.NumberColumn("발주수량", min_value=0)
                },
                hide_index=True
            )
            
            st.divider()
            
            # 하단 컨트롤러
            c1, c2, c3 = st.columns([1, 1, 1])
            add_val = c1.number_input("➕ 추가 수량", value=1, min_value=1, key="dong_add_val")
            
            if c2.button("🚀 선택 상품 수량 더하기", use_container_width=True):
                # 에디터에서 선택된 인덱스 찾기
                # 실제 세션 데이터에 반영
                for i, row in edited_df.iterrows():
                    if row['선택']:
                        # 원본 데이터의 인덱스를 찾아 업데이트
                        st.session_state.df_dong_current.at[i, '발주수량'] += add_val
                st.success(f"선택 항목에 {add_val}개씩 추가되었습니다.")
                st.rerun()
            
            # 다운로드 버튼
            csv_dong = edited_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            c3.download_button(
                label="📥 사입 리스트 다운로드 (CSV)", 
                data=csv_dong, 
                file_name=f"동대문사입_{datetime.now().strftime('%m%d')}.csv", 
                mime="text/csv",
                use_container_width=True
            )
    else:
        st.info("💡 동대문 발주용 엑셀 파일을 업로드해주세요.")
