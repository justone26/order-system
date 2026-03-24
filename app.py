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
        hist_date_4 = f_c3.date_input("🗓️ 입고 매핑 날짜", datetime.now(), key="date_v4_input_final_v2")

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
                edits = st.session_state["editor_v4_reorder_fix"].get("edited_rows", {})
                if edits:
                    for r_idx_str, change in edits.items():
                        orig_idx = df_work.index[int(r_idx_str)]
                        
                        # 1) 리오더 수량 직접 수정 시 반영
                        if "리오더 수량" in change:
                            st.session_state.df_raw.at[orig_idx, "리오더 수량"] = int(change["리오더 수량"])
                        
                        # 2) 💡 리오더 입고수량 입력 시 -> 오직 '리오더 수량'에서만 차감
                        if "리오더 입고수량" in change:
                            in_qty = int(change["리오더 입고수량"])
                            if in_qty > 0:
                                current_reorder = int(st.session_state.df_raw.at[orig_idx, "리오더 수량"])
                                # 리오더 수량만 차감 (가용재고는 변경 안 함)
                                st.session_state.df_raw.at[orig_idx, "리오더 수량"] = max(0, current_reorder - in_qty)
                                
                                # 입고 히스토리 저장 (추후 대조용)
                                log_df = pd.DataFrame([[df_work.at[orig_idx, item], df_work.at[orig_idx, option], in_qty]], columns=['상품명', '옵션', '수량'])
                                save_history_to_gsheet(log_df, log_type="입고")

                    # 최종 저장 및 화면 갱신
                    save_reorder_data(st.session_state.df_raw[[item, option, '리오더 수량']].rename(columns={item:'상품명', option:'옵션'}))
                    st.success("✅ 리오더 수량 차감 및 저장이 완료되었습니다!")
                    st.rerun()
                    
# --- [5단계: 최종 발주 리스트 요약 - 8개 컬럼 및 리오더 실시간 연동 완결판] ---
        st.divider()
        st.subheader("📋 5단계: 최종 발주 리스트 요약")
        
        # 1. 추가발주 입력값 임시 보관용 (입력창 초기화 방지)
        if 'add_order_dict' not in st.session_state: 
            st.session_state.add_order_dict = {}

        # 💡 [연동] 4단계와 동일한 최신 원본 데이터를 실시간으로 가져옵니다.
        df_5 = st.session_state.df_raw.copy()

        # 숫자형 변환 (계산 오류 및 16,000장 버그 차단)
        num_cols_5 = [avail, '리오더 수량', t7day, t3day]
        for c in num_cols_5:
            if c in df_5.columns:
                df_5[c] = pd.to_numeric(df_5[c], errors='coerce').fillna(0).astype(int)

        # 2. [계산 로직] 일판매량 반올림 및 권장발주량
        v7_5 = df_5[t7day]; v3_5 = df_5[t3day]
        # 일판매량: 반올림 후 정수 처리 (소수점 제거)
        df_5['일판매량'] = (v7_5 / 7 if v7_5.sum() > 0 else v3_5 / 3).round(0).astype(int)
        
        # 권장발주량 = (일판매량 * 확보일수) - (가용재고 + 리오더 수량)
        df_5['권장발주량'] = ((df_5['일판매량'] * (lt + ss)) - (df_5[avail] + df_5['리오더 수량'])).clip(lower=0).astype(int)
        
        # 추가발주수량 입력 컬럼 생성
        df_5['추가발주수량'] = df_5.index.map(st.session_state.add_order_dict).fillna(0).astype(int)

        # 상태 판별 (🚨긴급 / ⚠️주의 / ✅정상)
        def get_final_status(r):
            stock_sum = r[avail] + r['리오더 수량']
            daily = r['일판매량']
            if daily > 0:
                if stock_sum < (daily * 3): return "🚨 긴급"
                if stock_sum < (daily * 5): return "⚠️ 주의"
            return "✅ 정상"
        df_5['상태'] = df_5.apply(get_final_status, axis=1)

        # 3. [UI] 검색 및 필터링
        c5_1, c5_2, c5_3 = st.columns([1.5, 1.5, 1])
        search_q_v5 = c5_2.text_input("🔍 전체 상품명 검색 (입력 시 상태필터 무시)", key="v5_full_ordered_final")
        s_filter = c5_1.selectbox("🎯 상태 필터", ["🚨긴급 + ⚠️주의 우선", "🚨 긴급만 보기", "✅ 전체보기"], index=0)
        hist_date_5 = c5_3.date_input("🗓️ 기록 확인 날짜", datetime.now(), key="date_v5_final")

        # 필터링 적용 (검색 우선)
        if search_q_v5:
            df_5 = df_5[df_5[item].astype(str).str.contains(search_q_v5, case=False, na=False)]
        else:
            if s_filter == "🚨긴급 + ⚠️주의 우선": 
                df_5 = df_5[df_5['상태'].isin(["🚨 긴급", "⚠️ 주의"]) | (df_5['권장발주량'] > 0)]
            elif s_filter == "🚨 긴급만 보기": 
                df_5 = df_5[df_5['상태'] == "🚨 긴급"]
        
        df_5 = df_5.sort_values(by='상태')

        # 🎯 [순서 및 명칭 정리] 사장님 요청 8개 컬럼
        # 상태 => 상품명 => 옵션 => 공급쳐상품명 => 가용재고 => 리오더수량 => 추가발주수량 => 권장발주량
        df_display_5 = df_5.rename(columns={
            item: "상품명",
            option: "옵션",
            v_item: "공급쳐상품명",
            avail: "가용재고",
            "리오더 수량": "리오더수량"
        })

        final_cols_5 = [
            "상태", "상품명", "옵션", "공급쳐상품명", 
            "가용재고", "리오더수량", "추가발주수량", "권장발주량"
        ]
        actual_cols_5 = [c for c in final_cols_5 if c in df_display_5.columns]

        # 4. 데이터 에디터 및 발주 확정 (리오더 수량 합산 로직)
        with st.form("form_step_5_sync_and_save"):
            edited_v5 = st.data_editor(df_display_5[actual_cols_5], use_container_width=True, key="editor_v5_ordered_final", hide_index=True)
            submit_v5 = st.form_submit_button("✅ 추가발주 확정 및 리오더 수량 반영", use_container_width=True, type="primary")
            
            if submit_v5:
                edits = st.session_state["editor_v5_ordered_final"].get("edited_rows", {})
                if edits:
                    for r_idx_str, change in edits.items():
                        orig_idx = df_5.index[int(r_idx_str)]
                        
                        if "추가발주수량" in change:
                            add_qty = int(change["추가발주수량"])
                            if add_qty > 0:
                                # 💡 [연동] 추가 발주한 만큼 원본 리오더 수량에 더하기
                                current_re = int(st.session_state.df_raw.at[orig_idx, "리오더 수량"])
                                st.session_state.df_raw.at[orig_idx, "리오더 수량"] = current_re + add_qty
                                # 입력창 데이터 유지용
                                st.session_state.add_order_dict[orig_idx] = add_qty
                    
                    # 리오더 수량 데이터 구글 시트 백업 (필요 시 주석 해제)
                    # save_reorder_data(st.session_state.df_raw[[item, option, '리오더 수량']].rename(columns={item:'상품명', option:'옵션'}))
                    
                    st.success("✅ 발주량이 리오더 수량에 반영되었습니다. 4단계에서도 확인 가능합니다!")
                    st.rerun()

        st.write("---")
        
        # 5. 📥 [사장님 질문] 엑셀 다운로드 버튼 (현재 보이는 리스트 기준)
        if not df_display_5.empty:
            csv_data = df_display_5[actual_cols_5].to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button(
                label="📥 현재 발주 리스트 엑셀(CSV) 다운로드",
                data=csv_data,
                file_name=f"최종발주서_{datetime.now().strftime('%m%d_%H%M')}.csv",
                mime="text/csv",
                use_container_width=True
            )

        
# --- [6단계: 히스토리 조회 - 공급쳐 상품명 대응] ---
        if st.button("🔍 히스토리 조회하기", use_container_width=True, type="secondary"):
            raw_h = load_history_from_gsheet()
            if not raw_h.empty:
                raw_h['저장시간'] = pd.to_datetime(raw_h['저장시간'], errors='coerce')
                raw_h['pure_date'] = raw_h['저장시간'].dt.date
                f_df = raw_h[(raw_h['pure_date'] >= start_date) & (raw_h['pure_date'] <= end_date)].copy()

                if not f_df.empty:
                    # 데이터 이름 매칭
                    if '최종발주량' in f_df.columns: f_df['리오더 수량'] = f_df['최종발주량']
                    elif '수량' in f_df.columns: f_df['리오더 수량'] = f_df['수량']
                    
                    # 🎯 '공급쳐 상품명' 등 누락된 칸 보충
                    for col in ['공급쳐 상품명', '추가 리오더', '권장 발주량']:
                        if col not in f_df.columns: f_df[col] = 0 if '량' in col else ""

                    # 순서 고정
                    final_order = ['저장시간', '구분', '상품명', '옵션', '공급쳐 상품명', '리오더 수량', '추가 리오더', '권장 발주량']
                    actual_view = [c for c in final_order if c in f_df.columns]
                    show_df = f_df[actual_view].sort_values(by='저장시간', ascending=False)

                    st.success("✅ 조회가 완료되었습니다.")
                    st.dataframe(show_df, use_container_width=True, hide_index=True)
                    
                    csv_res = show_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
                    st.download_button("📥 엑셀 다운로드", csv_res, f"히스토리.csv", use_container_width=True)

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
