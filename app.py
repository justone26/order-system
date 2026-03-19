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

# 4단계: 편집 (품절부터 권장발주량까지 요청하신 순서 완벽 고정)
        if st.session_state.analyzed:
            st.subheader("📊 4단계: 데이터 편집 및 재고 관리")
            
            # --- [필터 및 검색] ---
            f1, f2 = st.columns([3, 1])
            search_query = f1.text_input("🔍 상품명 검색", key="prod_search_input")
            filter_mode = f2.selectbox("품절 필터", ["전체보기", "정상만", "품절만"], index=1)
            
            df_working = st.session_state.df_raw.copy()

            # --- [데이터 전처리 및 계산] ---
            # 숫자 변환 및 결측치 0 처리
            v_avail = pd.to_numeric(df_working[avail], errors='coerce').fillna(0)
            v_reorder = pd.to_numeric(df_working['리오더 수량'], errors='coerce').fillna(0)
            v_3day = pd.to_numeric(df_working[t3day], errors='coerce').fillna(0)
            
            # 1. 일판매량 계산 (3일 합계 / 3 -> 반올림 후 정수)
            df_working["일판매량"] = (v_3day / 3).round(0).astype(int)
            
            # 2. 권장발주량 계산 공식
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

            # --- [자동 저장 및 입고 차감 함수] ---
            def auto_save_and_update():
                if "main_editor" in st.session_state and st.session_state["main_editor"]["edited_rows"]:
                    changes = st.session_state["main_editor"]["edited_rows"]
                    for row_idx_str, change in changes.items():
                        row_idx = int(row_idx_str)
                        orig_idx = df_working.index[row_idx]
                        
                        if "리오더 수량" in change:
                            st.session_state.df_raw.at[orig_idx, "리오더 수량"] = int(round(float(change["리오더 수량"])))
                        
                        if "리오더입고수량" in change:
                            in_qty = int(round(float(change["리오더입고수량"])))
                            current_reorder = st.session_state.df_raw.at[orig_idx, "리오더 수량"]
                            st.session_state.df_raw.at[orig_idx, "리오더 수량"] = max(0, current_reorder - in_qty)
                    
                    try:
                        save_df = st.session_state.df_raw[[item, option, '리오더 수량']].copy()
                        save_df.columns = ['상품명', '옵션', '리오더 수량']
                        save_reorder_data(save_df)
                        st.toast("✅ 구글 시트 자동 저장 완료!")
                    except Exception as e:
                        st.error(f"저장 오류: {e}")

            # --- [화면 출력 설정: 요청하신 11개 컬럼 순서 고정] ---
            display_cols = [
                sold_out,    # 품절
                item,        # 상품명
                option,      # 옵션
                vendor_item, # 공급처상품명
                stock,       # 정상재고
                avail,       # 가용재고
                "리오더 수량",
                "리오더입고수량",
                "일판매량",
                t3day,       # 3일발주합계
                "권장발주량"
            ]
            
            # 리오더입고수량 칸 생성
            if "리오더입고수량" not in df_working.columns:
                df_working["리오더입고수량"] = 0

            # 실제 존재하는 컬럼만 필터링 (에러 방지)
            final_target = [c for c in display_cols if c in df_working.columns or c in ["일판매량", "권장발주량", "리오더입고수량"]]

            st.data_editor(
                df_working[final_target],
                use_container_width=True,
                key="main_editor",
                on_change=auto_save_and_update,
                column_config={
                    sold_out: st.column_config.Column("❌ 품절여부", disabled=True),
                    item: st.column_config.Column("📦 상품명", disabled=True),
                    option: st.column_config.Column("🎨 옵션", disabled=True),
                    vendor_item: st.column_config.Column("🏢 공급처상품명", disabled=True),
                    stock: st.column_config.NumberColumn("🔢 정상재고", disabled=True),
                    avail: st.column_config.NumberColumn("✅ 가용재고", disabled=True),
                    "리오더 수량": st.column_config.NumberColumn("📝 리오더 수량"),
                    "리오더입고수량": st.column_config.NumberColumn("➕ 입고수량 입력"),
                    "일판매량": st.column_config.NumberColumn("📈 일판매량", disabled=True),
                    t3day: st.column_config.NumberColumn("📊 3일발주합계", disabled=True),
                    "권장발주량": st.column_config.NumberColumn("🚀 권장발주량", disabled=True)
                }
            )
# 5단계: 요약 및 저장 (저장 및 다운로드 버튼 분리)
            st.subheader("📋 5단계: 최종 발주 리스트 요약")
            
            if 'df_raw' in st.session_state:
                to_order = st.session_state.df_raw.copy()
                
                # 기초 계산 및 컬럼 정리
                v_3day_val = pd.to_numeric(to_order[t3day], errors='coerce').fillna(0)
                to_order['일판매량'] = (v_3day_val / 3).round(0).astype(int)
                
                if '권장 발주량' in to_order.columns:
                    to_order['권장발주량'] = to_order['권장 발주량']
                elif '권장발주량' not in to_order.columns:
                    to_order['권장발주량'] = 0

                # --- [긴급도 계산 함수] ---
                def check_urgency(row):
                    v_av = pd.to_numeric(row.get(avail, 0), errors='coerce') or 0
                    v_sl = pd.to_numeric(row.get('일판매량', 0), errors='coerce') or 0
                    v_re = pd.to_numeric(row.get('리오더 수량', 0), errors='coerce') or 0
                    if v_sl > 0 and (v_av + v_re) < (v_sl * 3):
                        return "🚨 긴급(품절위험)"
                    elif v_sl > 0 and (v_av + v_re) < (v_sl * 5):
                        return "⚠️ 주의"
                    return "✅ 정상"

                to_order['상태'] = to_order.apply(check_urgency, axis=1)

                # --- [상태 필터 UI] ---
                f_col1, f_col2 = st.columns([2, 2])
                status_filter = f_col1.selectbox(
                    "🎯 상태 필터 선택", 
                    ["전체보기", "🚨 긴급만 보기", "⚠️ 주의이상 보기"],
                    key="final_status_filter_v2"
                )

                # 발주 필요건 필터링
                mask = (pd.to_numeric(to_order['권장발주량'], errors='coerce') > 0) | (to_order['상태'] != "✅ 정상")
                to_order = to_order[mask].copy()

                if status_filter == "🚨 긴급만 보기":
                    to_order = to_order[to_order['상태'] == "🚨 긴급(품절위험)"]
                elif status_filter == "⚠️ 주의이상 보기":
                    to_order = to_order[to_order['상태'].str.contains("🚨|⚠️")]

                if not to_order.empty:
                    if '추가발주분' not in to_order.columns:
                        to_order['추가발주분'] = 0
                    
                    # 8개 컬럼 순서 고정
                    final_display_cols = ["상태", item, option, vendor_item, avail, "리오더 수량", "추가발주분", "권장발주량"]
                    final_display_cols = [c for c in final_display_cols if c in to_order.columns or c in ["상태", "추가발주분"]]

                    st.write(f"### ✏️ 발주 수량 확인 ({status_filter})")
                    
                    # 데이터 편집기
                    final_order_df = st.data_editor(
                        to_order[final_display_cols], 
                        use_container_width=True, 
                        key="final_order_editor_v2",
                        column_config={
                            "상태": st.column_config.Column("📢 상태", width="medium"),
                            "추가발주분": st.column_config.NumberColumn("➕ 추가발주분", help="이번에 새로 주문할 수량을 입력하세요."),
                            avail: st.column_config.NumberColumn("✅ 가용재고", disabled=True),
                            "리오더 수량": st.column_config.NumberColumn("📦 현재리오더수량", disabled=True),
                            "권장발주량": st.column_config.NumberColumn("🚀 권장발주량", disabled=True)
                        }
                    )
                    
                    # --- [하단 버튼 배치] ---
                    st.write("") # 간격 조절
                    btn_col1, btn_col2 = st.columns(2)
                    
                    with btn_col1:
                        # 1. 구글 시트 저장 버튼
                        if st.button("💾 구글 시트에 데이터 저장", use_container_width=True, type="primary"):
                            for idx, row in final_order_df.iterrows():
                                current_re = int(row.get('리오더 수량', 0))
                                added_re = int(row.get('추가발주분', 0))
                                st.session_state.df_raw.at[idx, '리오더 수량'] = current_re + added_re
                            
                            save_df = st.session_state.df_raw[[item, option, '리오더 수량']].copy()
                            save_df.columns = ['상품명', '옵션', '리오더 수량']
                            save_reorder_data(save_df)
                            st.success("✅ 구글 시트에 업데이트가 완료되었습니다!")
                            st.rerun()
                    
                    with btn_col2:
                        # 2. 독립적인 엑셀(CSV) 다운로드 버튼
                        csv_data = final_order_df.to_csv(index=False).encode('utf-8-sig')
                        st.download_button(
                            label="📥 현재 화면 엑셀(CSV) 다운로드",
                            data=csv_data,
                            file_name=f"발주확인서_{datetime.now().strftime('%m%d_%H%M')}.csv",
                            mime="text/csv",
                            use_container_width=True
                        )
                else:
                    st.info(f"필터 조건({status_filter})에 맞는 상품이 없습니다.")
                        
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
