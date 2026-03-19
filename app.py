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
            
            # 구글 시트 연동 로직 (리오더 수량 복구)
            try:
                sheet = get_sheet().sheet1
                gs_data = pd.DataFrame(sheet.get_all_records())
                
                if not gs_data.empty and '상품명' in gs_data.columns:
                    # 매칭을 위해 상품명, 옵션, 리오더 수량만 가져옴
                    gs_subset = gs_data[['상품명', '옵션', '리오더 수량']].copy()
                    # 기존 데이터와 합치기 (상품명+옵션 기준)
                    df_new = pd.merge(df_new, gs_subset, on=['상품명', '옵션'], how='left', suffixes=('', '_gs'))
                    
                    if '리오더 수량_gs' in df_new.columns:
                        df_new['리오더 수량'] = df_new['리오더 수량_gs'].fillna(0).astype(int)
                        df_new = df_new.drop(columns=['리오더 수량_gs'])
                    else:
                        if '리오더 수량' not in df_new.columns: df_new['리오더 수량'] = 0
                else:
                    if '리오더 수량' not in df_new.columns: df_new['리오더 수량'] = 0
            except Exception as e:
                st.warning(f"기존 리오더 수량 불러오기 실패: {e}")
                if '리오더 수량' not in df_new.columns: df_new['리오더 수량'] = 0

            st.session_state.df_raw = df_new
            st.session_state.last_filename = uploaded_file.name
            st.session_state.analyzed = False
            st.rerun()

    # 데이터가 로드된 경우에만 매핑 및 분석 진행
    if st.session_state.get('df_raw') is not None:
        df_current = st.session_state.df_raw
        cols = df_current.columns.tolist()

        # 1단계: 매핑 설정
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

        # 2~3단계: 분석 설정
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

        # 4단계: 편집 (숫자 정렬 꼼수 적용)
        if st.session_state.analyzed:
            st.subheader("📊 4단계: 데이터 편집 및 재고 관리")
            
            f1, f2 = st.columns([3, 1])
            search_query = f1.text_input("🔍 상품명 검색", key="prod_search_input")
            filter_mode = f2.selectbox("품절 필터", ["전체보기", "정상만", "품절만"], index=1)
            
            df_working = st.session_state.df_raw.copy()

            v_avail = pd.to_numeric(df_working[avail], errors='coerce').fillna(0)
            v_reorder = pd.to_numeric(df_working['리오더 수량'], errors='coerce').fillna(0)
            v_3day = pd.to_numeric(df_working[t3day], errors='coerce').fillna(0)
            
            df_working["일판매량"] = (v_3day / 3).round(0).astype(int)
            needed_qty = (df_working["일판매량"] * (lead_time + safety_stock))
            current_assets = (v_avail + v_reorder)
            df_working["권장발주량"] = (needed_qty - current_assets).clip(lower=0).round(0).astype(int)

            if filter_mode == "정상만": 
                df_working = df_working[~df_working[sold_out].astype(str).str.contains('품절', na=False)]
            elif filter_mode == "품절만": 
                df_working = df_working[df_working[sold_out].astype(str).str.contains('품절', na=False)]
            if search_query: 
                df_working = df_working[df_working[item].astype(str).str.contains(search_query, case=False, na=False)]

            if "리오더입고수량" not in df_working.columns:
                df_working["리오더입고수량"] = 0

            # 자동 저장 함수
            def auto_save_and_update():
                if "main_editor" in st.session_state and st.session_state["main_editor"]["edited_rows"]:
                    changes = st.session_state["main_editor"]["edited_rows"]
                    for row_idx_str, change in changes.items():
                        row_idx = int(row_idx_str)
                        orig_idx = df_working.index[row_idx]
                        
                        if "리오더 수량" in change:
                            # 꼼수로 들어간 공백 제거 후 숫자로 저장
                            val = str(change["리오더 수량"]).replace(" ", "")
                            st.session_state.df_raw.at[orig_idx, "리오더 수량"] = int(float(val))
                        
                        if "리오더입고수량" in change:
                            val = str(change["리오더입고수량"]).replace(" ", "")
                            in_qty = int(float(val))
                            current_reorder = st.session_state.df_raw.at[orig_idx, "리오더 수량"]
                            st.session_state.df_raw.at[orig_idx, "리오더 수량"] = max(0, current_reorder - in_qty)
                    
                    try:
                        save_df = st.session_state.df_raw[[item, option, '리오더 수량']].copy()
                        save_df.columns = ['상품명', '옵션', '리오더 수량']
                        save_reorder_data(save_df)
                        st.toast("✅ 구글 시트 자동 저장 완료!")
                    except Exception as e:
                        st.error(f"저장 오류: {e}")

            # [꼼수 적용] 표시용 데이터 가공
            display_df_4 = df_working.copy()
            num_cols_4 = [stock, avail, "리오더 수량", "리오더입고수량", "일판매량", t3day, "권장발주량"]
            for col in num_cols_4:
                if col in display_df_4.columns:
                    display_df_4[col] = display_df_4[col].fillna(0).astype(int).apply(lambda x: f"  {x}")

            display_cols_4 = [sold_out, item, option, vendor_item, stock, avail, "리오더 수량", "리오더입고수량", "일판매량", t3day, "권장발주량"]
            final_target_4 = [c for c in display_cols_4 if c in display_df_4.columns]

            st.data_editor(
                display_df_4[final_target_4],
                use_container_width=True,
                height=600, 
                key="main_editor",
                on_change=auto_save_and_update,
                column_config={
                    sold_out: st.column_config.Column("❌ 품절여부", disabled=True),
                    item: st.column_config.Column("📦 상품명", disabled=True, width="medium"),
                    option: st.column_config.Column("🎨 옵션", disabled=True, width="small"),
                    vendor_item: st.column_config.Column("🏢 공급처상품명", disabled=True, width=280),
                    stock: st.column_config.Column("🔢 정상재고", disabled=True),
                    avail: st.column_config.Column("✅ 가용재고", disabled=True),
                    "리오더 수량": st.column_config.Column("📝 리오더 수량"),
                    "리오더입고수량": st.column_config.Column("➕ 입고수량 입력"),
                    "일판매량": st.column_config.Column("📈 일판매량", disabled=True),
                    t3day: st.column_config.Column("📊 3일발주합계", disabled=True),
                    "권장발주량": st.column_config.Column("🚀 권장발주량", disabled=True)
                }
            )

   # 5단계 편집기 바로 아래에 위치해야 함     
                    st.divider()
                    col_btn1, col_btn2 = st.columns(2)

                    # 1. 구글 시트 'history' 탭에 저장 (과거 기록용)
                    if col_btn1.button("💾 구글 시트에 최종 기록 저장"):
                        # 표시용 공백 제거 후 순수 데이터만 추출
                        save_data = final_order_df.copy()
                        for c in save_data.columns:
                            if save_data[c].dtype == object:
                                save_data[c] = save_data[c].astype(str).str.strip()
                        
                        success = save_history_to_gsheet(save_data)
                        if success:
                            st.success("✅ 'history' 탭에 오늘 기록이 추가되었습니다!")
                        else:
                            st.error("❌ 저장 실패. 설정을 확인하세요.")

                    # 2. 엑셀 파일로 다운로드
                    download_df = final_order_df.copy()
                    for c in download_df.columns:
                        if download_df[c].dtype == object:
                            download_df[c] = download_df[c].astype(str).str.strip()

                    csv = download_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
                    col_btn2.download_button(
                        label="📥 발주 리스트 엑셀 다운로드",
                        data=csv,
                        file_name=f"발주리스트_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                        mime="text/csv"
                    )
                # 이 else는 if not to_order.empty: 와 짝궁입니다.
                else:
                    st.info("💡 발주할 상품이 없습니다. (권장발주량 > 0 또는 긴급/주의 상품 없음)")
                    
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
