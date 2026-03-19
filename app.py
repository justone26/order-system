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

# 4단계: 편집 (숫자 정렬 꼼수 및 표 크기 최적화 적용)
        if st.session_state.analyzed:
            st.subheader("📊 4단계: 데이터 편집 및 재고 관리")
            
            # --- [필터 및 검색] ---
            f1, f2 = st.columns([3, 1])
            search_query = f1.text_input("🔍 상품명 검색", key="prod_search_input")
            filter_mode = f2.selectbox("품절 필터", ["전체보기", "정상만", "품절만"], index=1)
            
            df_working = st.session_state.df_raw.copy()

            # --- [데이터 전처리 및 계산] ---
            v_avail = pd.to_numeric(df_working[avail], errors='coerce').fillna(0)
            v_reorder = pd.to_numeric(df_working['리오더 수량'], errors='coerce').fillna(0)
            v_3day = pd.to_numeric(df_working[t3day], errors='coerce').fillna(0)
            
            df_working["일판매량"] = (v_3day / 3).round(0).astype(int)
            
            needed_qty = (df_working["일판매량"] * (lead_time + safety_stock))
            current_assets = (v_avail + v_reorder)
            df_working["권장발주량"] = (needed_qty - current_assets).clip(lower=0).round(0).astype(int)

            # 필터 적용
            if filter_mode == "정상만": 
                df_working = df_working[~df_working[sold_out].astype(str).str.contains('품절', na=False)]
            elif filter_mode == "품절만": 
                df_working = df_working[df_working[sold_out].astype(str).str.contains('품절', na=False)]
            if search_query: 
                df_working = df_working[df_working[item].astype(str).str.contains(search_query, case=False, na=False)]

            # 리오더입고수량 칸 생성
            if "리오더입고수량" not in df_working.columns:
                df_working["리오더입고수량"] = 0

            # --- [자동 저장 함수 유지] ---
            def auto_save_and_update():
                if "main_editor" in st.session_state and st.session_state["main_editor"]["edited_rows"]:
                    changes = st.session_state["main_editor"]["edited_rows"]
                    for row_idx_str, change in changes.items():
                        row_idx = int(row_idx_str)
                        orig_idx = df_working.index[row_idx]
                        
                        # 수정 시 다시 숫자로 변환하여 저장 (중요!)
                        if "리오더 수량" in change:
                            st.session_state.df_raw.at[orig_idx, "리오더 수량"] = int(round(float(str(change["리오더 수량"]).strip())))
                        
                        if "리오더입고수량" in change:
                            in_qty = int(round(float(str(change["리오더입고수량"]).strip())))
                            current_reorder = st.session_state.df_raw.at[orig_idx, "리오더 수량"]
                            st.session_state.df_raw.at[orig_idx, "리오더 수량"] = max(0, current_reorder - in_qty)
                    
                    try:
                        save_df = st.session_state.df_raw[[item, option, '리오더 수량']].copy()
                        save_df.columns = ['상품명', '옵션', '리오더 수량']
                        save_reorder_data(save_df)
                        st.toast("✅ 구글 시트 자동 저장 완료!")
                    except Exception as e:
                        st.error(f"저장 오류: {e}")

            # --- [표시용 데이터 가공: 숫자 왼쪽 정렬 꼼수] ---
            display_df = df_working.copy()
            
            # 공백을 추가하여 왼쪽으로 밀어낼 숫자 컬럼 리스트
            num_cols_to_fix = [stock, avail, "리오더 수량", "리오더입고수량", "일판매량", t3day, "권장발주량"]
            
            for col in num_cols_to_fix:
                if col in display_df.columns:
                    # 숫자를 문자로 바꾸고 앞뒤에 공백을 넣어 왼쪽 정렬 느낌 유도
                    display_df[col] = display_df[col].fillna(0).astype(int).apply(lambda x: f"  {x}")

            # 컬럼 순서 고정
            display_cols = [sold_out, item, option, vendor_item, stock, avail, "리오더 수량", "리오더입고수량", "일판매량", t3day, "권장발주량"]
            final_target = [c for c in display_cols if c in display_df.columns]

            # --- [최종 화면 출력] ---
            st.data_editor(
                display_df[final_target],
                use_container_width=True,
                height=600,  # 👈 높이를 600으로 시원하게 키웠습니다.
                key="main_editor",
                on_change=auto_save_and_update,
                column_config={
                    sold_out: st.column_config.Column("❌ 품절여부", disabled=True),
                    item: st.column_config.Column("📦 상품명", disabled=True, width="medium"),
                    option: st.column_config.Column("🎨 옵션", disabled=True, width="small"),
                    vendor_item: st.column_config.Column("🏢 공급처상품명", disabled=True, width=280),
                    # [주의] Column으로 설정해야 왼쪽 정렬(꼼수)이 유지됩니다.
                    stock: st.column_config.Column("🔢 정상재고", disabled=True),
                    avail: st.column_config.Column("✅ 가용재고", disabled=True),
                    "리오더 수량": st.column_config.Column("📝 리오더 수량"),
                    "리오더입고수량": st.column_config.Column("➕ 입고수량 입력"),
                    "일판매량": st.column_config.Column("📈 일판매량", disabled=True),
                    t3day: st.column_config.Column("📊 3일발주합계", disabled=True),
                    "권장발주량": st.column_config.Column("🚀 권장발주량", disabled=True)
                }
            )
            
# 5단계: 요약 및 저장 (KeyError 해결 및 안전한 컬럼 매칭)
            st.subheader("📋 5단계: 최종 발주 리스트 요약")
            
            if 'df_raw' in st.session_state:
                # 1. 계산용 데이터 준비
                to_order = st.session_state.df_raw.copy()
                
                # 수치 계산
                v_3day_val = pd.to_numeric(to_order[t3day], errors='coerce').fillna(0)
                to_order['일판매량'] = (v_3day_val / 3).round(0).astype(int)
                
                # 권장발주량 컬럼 확인 및 생성
                if '권장 발주량' in to_order.columns:
                    to_order['권장발주량'] = to_order['권장 발주량']
                elif '권장발주량' not in to_order.columns:
                    to_order['권장발주량'] = 0

                # 긴급도 계산
                def check_urgency(row):
                    v_av = pd.to_numeric(row.get(avail, 0), errors='coerce') or 0
                    v_sl = pd.to_numeric(row.get('일판매량', 0), errors='coerce') or 0
                    v_re = pd.to_numeric(row.get('리오더 수량', 0), errors='coerce') or 0
                    if v_sl > 0 and (v_av + v_re) < (v_sl * 3): return "🚨 긴급"
                    elif v_sl > 0 and (v_av + v_re) < (v_sl * 5): return "⚠️ 주의"
                    return "✅ 정상"

                to_order['상태'] = to_order.apply(check_urgency, axis=1)

                # 필터링 로직 (생략 - 기존과 동일)
                # ...

                if not to_order.empty:
                    # [꼼수] 표시용 데이터 생성 (숫자를 문자로 변환하여 우측 정렬 강제 해제)
                    display_df = to_order.copy()
                    if '추가발주분' not in display_df.columns:
                        display_df['추가발주분'] = 0

                    # 2. [에러 방지] 실제 존재하는 컬럼만 리스트업
                    # vendor_item 변수가 비어있을 경우를 대비해 직접 이름을 확인해
                    potential_cols = ["상태", item, option, vendor_item, avail, "리오더 수량", "추가발주분", "권장발주량"]
                    
                    # 실제로 display_df에 존재하는 컬럼만 필터링 (KeyError 방지 핵심!)
                    existing_cols = [c for c in potential_cols if c in display_df.columns]
                    
                    # 숫자 컬럼들을 문자로 변환 (중앙 정렬 느낌 유도)
                    num_targets = [avail, "리오더 수량", "추가발주분", "권장발주량"]
                    for col in num_targets:
                        if col in display_df.columns:
                            display_df[col] = display_df[col].fillna(0).astype(int).astype(str)

                    st.write(f"### ✏️ 발주 수량 확인")
                    
                    # 3. 데이터 편집기 실행
                    final_order_df = st.data_editor(
                        display_df[existing_cols], # 안전하게 존재하는 컬럼만 넣음!
                        use_container_width=True, 
                        key="final_order_editor_no_keyerror",
                        column_config={
                            "상태": st.column_config.Column("📢 상태", width=70),
                            item: st.column_config.Column("📦 상품명", width="medium"), 
                            option: st.column_config.Column("🎨 옵션", width=80),
                            vendor_item: st.column_config.Column("🏢 공급처상품명", width=280), 
                            avail: st.column_config.Column("가용재고", width=80),
                            "리오더 수량": st.column_config.Column("현 리오더", width=80),
                            "추가발주분": st.column_config.Column("추가발주", width=80),
                            "권장발주량": st.column_config.Column("권장발주", width=80)
                        }
                    )
                    
                    st.divider()
                    # (이후 저장 버튼 로직은 row.get()을 사용하여 안전하게 처리)
              
            # 6단계: 과거 확인
            st.subheader("📜 6단계: 과거 데이터 확인")
            if st.button("🔄 기록 불러오기"):
                st.session_state.db_history = load_history_from_gsheet()
            if 'db_history' in st.session_state and not st.session_state.db_history.empty:
                df_hist = st.session_state.db_history
                df_hist['날짜'] = df_hist['저장시간'].astype(str).str.split(' ').str[0]
                sel_date = st.date_input("날짜 선택", datetime.now())
                target_date = sel_date.strftime("%Y-%m-%d")
                day_data = df_hist[df_hist['날짜'] == target_date]
                if not day_data.empty:
                    sel_time = st.selectbox("⏰ 시간 선택", sorted(day_data['저장시간'].unique(), reverse=True))
                    st.dataframe(day_data[day_data['저장시간'] == sel_time].drop(columns=['날짜']), use_container_width=True)
                else: st.info("📅 해당 날짜 기록 없음")

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
