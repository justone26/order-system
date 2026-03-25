import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 1. [핵심] 한국 표준시(KST) 및 날짜 설정
def get_now_kst():
    return datetime.now(timezone(timedelta(hours=9)))

now = get_now_kst()
today_date = now.strftime("%Y-%m-%d")
today_time = now.strftime("%H:%M:%S")

st.set_page_config(layout="wide", page_title=f"저스트원 재고관리 ({now.strftime('%m/%d')})")

# [나중에 쓸 전송용 함수]
def save_to_google(df_to_save):
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        sh = client.open_by_key("1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg")
        # 여기에 저장할 시트 이름(예: '발주현황')을 넣으면 됩니다.
        ws = sh.worksheet("Sheet1") 
        # 데이터프레임을 리스트로 변환하여 추가
        ws.append_rows(df_to_save.values.tolist())
        return True
    except:
        return False

# 리셋 콜백
def reset_callback():
    for key in list(st.session_state.keys()):
        del st.session_state[key]

# 자동 매칭 로직
def get_auto_index(cols, keywords):
    for i, col in enumerate(cols):
        if any(k in str(col).strip() for k in keywords):
            return i
    return 0

# 세션 초기화
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'df_final' not in st.session_state: st.session_state.df_final = None
if 'mapping' not in st.session_state: st.session_state.mapping = {}

# --- 화면 시작 ---
st.title("📦 저스트원 통합 재고 관리 시스템")
st.info(f"📅 **분석 기준일:** {today_date} / **현재 시간:** {today_time}")

tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 관리"])

with tab1:
    # 1단계: 업로드
    st.subheader("📁 1단계: 데이터 업로드")
    up_file = st.file_uploader("파일을 선택하세요", type=['xlsx', 'xls', 'csv'], key="main_uploader", label_visibility="collapsed")
    
    col_btn, _ = st.columns([1, 3])
    with col_btn:
        st.button("🔄 전체 데이터 초기화", use_container_width=True, on_click=reset_callback)

    if up_file is not None and st.session_state.df_raw is None:
        try:
            df = pd.read_csv(up_file) if up_file.name.endswith('.csv') else pd.read_excel(up_file)
            df.columns = df.columns.str.strip()
            st.session_state.df_raw = df
            st.rerun()
        except Exception as e:
            st.error(f"파일 읽기 실패: {e}")

    # 2단계 & 3단계: 설정 영역 (파일 있을 때 상시 노출)
    if st.session_state.df_raw is not None:
        st.divider()
        st.subheader("🔗 2단계: 자동 컬럼 매칭")
        cols = st.session_state.df_raw.columns.tolist()
        
        c1, c2 = st.columns(2)
        with c1:
            sold_out = st.selectbox("품절 여부", cols, index=get_auto_index(cols, ['품절', '판매중단']))
            vendor = st.selectbox("공급처", cols, index=get_auto_index(cols, ['공급처', '업체명']))
            item = st.selectbox("상품명", cols, index=get_auto_index(cols, ['상품명', '상품']))
            option = st.selectbox("옵션", cols, index=get_auto_index(cols, ['옵션']))
            vendor_item = st.selectbox("공급처 상품명", cols, index=get_auto_index(cols, ['공급처상품명', '거래처옵션']))
        with c2:
            reg_date = st.selectbox("등록일", cols, index=get_auto_index(cols, ['등록일', '생성일']))
            stock = st.selectbox("정상재고", cols, index=get_auto_index(cols, ['정상재고', '재고']))
            avail = st.selectbox("가용재고", cols, index=get_auto_index(cols, ['가용재고', '가용']))
            t3day = st.selectbox("3일 발주합계", cols, index=get_auto_index(cols, ['3일']))
            t1week = st.selectbox("7일 발주합계", cols, index=get_auto_index(cols, ['7일', '1주']))

        st.divider()
        st.subheader("⚙️ 3단계: 발주 기준 설정")
        col_lt, col_ss = st.columns(2)
        with col_lt:
            lt_val = st.number_input("⏳ 리드타임 (일) - 디폴트 7일", min_value=1, value=7)
        with col_ss:
            ss_val = st.number_input("🛡️ 안전재고 (일) - 디폴트 3일", min_value=0, value=3)

        if st.button("🚀 데이터 분석 시작", use_container_width=True):
            st.session_state.mapping = {
                "item": item, "option": option, "avail": avail, 
                "t1week": t1week, "lt": lt_val, "ss": ss_val
            }
            # 계산 로직
            df_calc = st.session_state.df_raw.copy()
            df_calc['일판매량'] = (pd.to_numeric(df_calc[t1week], errors='coerce').fillna(0) / 7).round(2)
            df_calc['필요재고'] = (df_calc['일판매량'] * (lt_val + ss_val)).round(0).astype(int)
            df_calc['가용재고_num'] = pd.to_numeric(df_calc[avail], errors='coerce').fillna(0)
            df_calc['권장발주량'] = (df_calc['필요재고'] - df_calc['가용재고_num']).clip(lower=0).astype(int)
            
            st.session_state.df_final = df_calc
            st.rerun()

  # ==========================================
# 4단계: 데이터 편집 및 재고 관리
# ==========================================
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📊 4단계: 데이터 편집 및 재고 관리")

    # 2단계에서 매칭된 컬럼 정보 가져오기
    m = st.session_state.mapping
    sold_out, vendor, v_item, item, option = m['sold_out'], m['vendor'], m['vendor_item'], m['item'], m['option']
    stock, avail, t3day, t1week = m['stock'], m['avail'], m['t3day'], m['t1week']
    lt, ss = m['lt'], m['ss']

    df_work = st.session_state.df_raw.copy()
    
    # 숫자형 변환 (에러 방지용)
    for c in [stock, avail, t1week, t3day]:
        if c in df_work.columns:
            df_work[c] = pd.to_numeric(df_work[c], errors='coerce').fillna(0).astype(int)
    
    # 리오더 관련 컬럼 초기화 및 변환
    if "리오더 수량" not in df_work.columns: df_work["리오더 수량"] = 0
    df_work["리오더 수량"] = pd.to_numeric(df_work["리오더 수량"], errors='coerce').fillna(0).astype(int)
    if "리오더 입고수량" not in df_work.columns: df_work["리오더 입고수량"] = 0

    # 5단계와 동일한 계산 로직 적용
    df_work['일판매량'] = df_work.apply(lambda x: round(x[t1week] / 7) if x[t1week] > 0 else round(x[t3day] / 3), axis=1).astype(int)
    df_work['권장발주량'] = ((df_work['일판매량'] * (lt + ss)) - (df_work[avail] + df_work['리오더 수량'])).clip(lower=0).astype(int)

    # UI 및 필터
    f_c1, f_c2, f_c3 = st.columns([2, 1, 1])
    search_q = f_c1.text_input("🔍 상품명/옵션 검색", key="v4_search")
    filter_m = f_c2.selectbox("상태 필터", ["전체보기", "정상만", "품절만"], index=1, key="v4_filter")
    hist_date_4 = f_c3.date_input("🗓️ 입고 기록 날짜", now.date(), key="v4_date")

    if filter_m == "정상만":
        df_work = df_work[~df_work[sold_out].astype(str).str.contains('품절', na=False)]
    elif filter_m == "품절만":
        df_work = df_work[df_work[sold_out].astype(str).str.contains('품절', na=False)]
    
    if search_q:
        df_work = df_work[df_work[item].astype(str).str.contains(search_q, case=False, na=False) | 
                          df_work[option].astype(str).str.contains(search_q, case=False, na=False)]

    display_cols = [sold_out, vendor, item, option, v_item, stock, avail, "리오더 수량", "리오더 입고수량", "일판매량", "권장발주량"]
    
    with st.form("form_v4_edit"):
        edited_v4 = st.data_editor(
            df_work[display_cols], 
            use_container_width=True, 
            key="ed_v4_work", 
            hide_index=True,
            column_config={
                "리오더 입고수량": st.column_config.NumberColumn("입고입력", help="입고된 수량을 입력하면 리오더 수량에서 차감됩니다.")
            }
        )
        if st.form_submit_button("💾 데이터 일시 저장 (앱 내 반영)", use_container_width=True, type="primary"):
            edits = st.session_state["ed_v4_work"].get("edited_rows", {})
            if edits:
                for r_idx, change in edits.items():
                    orig_idx = df_work.index[int(r_idx)]
                    # 입고 로직: 리오더수량 - 입고수량
                    if "리오더 입고수량" in change:
                        in_qty = int(change["리오더 입고수량"])
                        if in_qty > 0:
                            curr_reorder = int(st.session_state.df_raw.at[orig_idx, "리오더 수량"])
                            st.session_state.df_raw.at[orig_idx, "리오더 수량"] = max(0, curr_reorder - in_qty)
                    # 수량 직접 수정
                    if "리오더 수량" in change:
                        st.session_state.df_raw.at[orig_idx, "리오더 수량"] = int(change["리오더 수량"])
                
                st.success("✅ 변경사항이 반영되었습니다. 아래 5단계에서 최종 확인 후 저장하세요!")
                st.rerun()

# ==========================================
# 5단계: 최종 발주 및 히스토리 저장
# ==========================================
if st.session_state.get('analyzed') and st.session_state.df_raw is not None:
    st.divider()
    st.subheader("📋 5단계: 최종 발주 리스트 요약")

    df_5 = st.session_state.df_raw.copy()
    m = st.session_state.mapping
    
    # 상단 상태 판별 함수
    def get_status(r):
        tot = r[m['avail']] + r['리오더 수량']
        day = (r[m['t1week']] / 7)
        if day > 0:
            if tot < (day * 3): return "🚨 긴급"
            if tot < (day * 5): return "⚠️ 주의"
        return "✅ 정상"

    df_5['상태'] = df_5.apply(get_status, axis=1)
    df_5['일판매량'] = (df_5[m['t1week']] / 7).round(1)
    df_5['권장발주량'] = ((df_5['일판매량'] * (m['lt'] + m['ss'])) - (df_5[m['avail']] + df_5['리오더 수량'])).clip(lower=0).round(0).astype(int)

    final_display = ["상태", m['item'], m['option'], m['vendor_item'], m['avail'], "리오더 수량", "일판매량", "권장발주량"]
    
    st.dataframe(df_5[final_display], use_container_width=True, hide_index=True)

    col_save, col_csv = st.columns(2)
    with col_save:
        if st.button("💾 구글 시트에 최종 발주 기록 저장", use_container_width=True, type="primary"):
            # 발주량이 있는 데이터만 필터링
            to_save = df_5[df_5['권장발주량'] > 0].copy()
            if not to_save.empty:
                # 구글 시트 전송용 리스트 생성
                log_data = []
                for _, row in to_save.iterrows():
                    log_data.append([
                        today_date + " " + today_time,
                        row['상태'],
                        row[m['item']],
                        row[m['option']],
                        row[m['vendor_item']],
                        int(row[m['avail']]),
                        int(row['리오더 수량']),
                        int(row['권장발주량'])
                    ])
                
                # 시트 저장 실행
                if save_to_google(pd.DataFrame(log_data)):
                    st.success(f"🎉 {len(log_data)}건의 발주 내역이 구글 시트에 기록되었습니다!")
                else:
                    st.error("📡 구글 시트 저장 중 오류가 발생했습니다.")
            else:
                st.warning("발주 권장 수량이 있는 항목이 없습니다.")

    with col_csv:
        csv = df_5[final_display].to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
        st.download_button("📥 최종 발주서 CSV 다운로드", data=csv, file_name=f"JustOne_Order_{today_date}.csv", use_container_width=True)

# ==========================================
# 6단계: 히스토리 조회 (생략 가능 - 시트에서 직접 확인 권장)
# ==========================================
