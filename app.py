import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import holidays

# 페이지 설정
st.set_page_config(layout="wide", page_title="재고 관리 시스템")

# [1] 상태 관리 초기화
if 'file_key' not in st.session_state: st.session_state.file_key = 0
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'history' not in st.session_state: st.session_state.history = {}

st.title("📦 재고 관리 및 발주 시스템")

# [시스템 초기화]
if st.button("🔄 시스템 초기화 (데이터 및 파일 삭제)"):
    # history는 유지하고 싶다면 이 부분에서 제외해도 되지만, 완벽 초기화를 위해 전부 삭제함
    for key in list(st.session_state.keys()): del st.session_state[key]
    st.session_state.file_key = 1
    st.rerun()

# [파일 업로드]
uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'xls', 'csv'], key=f"uploader_{st.session_state.file_key}")

if uploaded_file is not None and st.session_state.df_raw is None:
    try:
        df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
        df = df.loc[:, ~df.columns.duplicated()]
        if "1차 리오더" not in df.columns: df["1차 리오더"] = 0
        if "2차 리오더" not in df.columns: df["2차 리오더"] = 0
        st.session_state.df_raw = df
        st.rerun()
    except Exception as e: st.error(f"파일 오류: {e}")

if st.session_state.df_raw is not None:
    def get_auto_index(cols, keywords):
        for key in keywords:
            for i, c in enumerate(cols):
                if key in str(c): return i
        return 0

    cols = st.session_state.df_raw.columns.tolist()

    # 1단계 매핑
    st.subheader("⚙️ 1단계: 자동 매핑 설정")
    c1, c2 = st.columns(2)
    with c1:
        sold_out = st.selectbox("품절 여부", cols, index=get_auto_index(cols, ['품절', '판매중단']))
        vendor = st.selectbox("공급처", cols, index=get_auto_index(cols, ['공급처', '업체명']))
        item = st.selectbox("상품명", cols, index=get_auto_index(cols, ['상품명', '상품']))
        option = st.selectbox("옵션", cols, index=get_auto_index(cols, ['옵션']))
        reg_date_col = st.selectbox("등록일 컬럼", cols, index=get_auto_index(cols, ['등록일', '생성일', '입점일']))
    with c2:
        vendor_item_name = st.selectbox("공급처 상품명", cols, index=get_auto_index(cols, ['공급처상품명', '거래처옵션', '공급처옵션']))
        stock = st.selectbox("정상재고", cols, index=get_auto_index(cols, ['정상재고', '재고']))
        avail = st.selectbox("가용재고", cols, index=get_auto_index(cols, ['가용재고', '가용']))
        t3day = st.selectbox("3일 발주 합계", cols, index=get_auto_index(cols, ['3일', '최근3일']))
        t1week = st.selectbox("1주 발주 합계", cols, index=get_auto_index(cols, ['1주', '7일', '최근7일']))

    # 2~3단계 분석
    st.subheader("⚙️ 2~3단계: 기간 설정 및 분석")
    l1, l2 = st.columns(2)
    lead_time = l1.number_input("리드타임 (일)", value=0)
    safety_stock = l2.number_input("안전재고 (일)", value=3)
    
    if st.button("🚀 분석 실행"):
        df = st.session_state.df_raw
        today = datetime.now()
        kr_holidays = holidays.KR(years=today.year)
        def get_biz_days(start_date):
            if pd.isna(start_date): return 3
            days = pd.date_range(start=start_date, end=today)
            biz_days = [d for d in days if d.weekday() < 5 and d not in kr_holidays]
            return max(1, len(biz_days))
        
        reg_dates = pd.to_datetime(df[reg_date_col], errors='coerce')
        divisors = [min(3, get_biz_days(rd)) for rd in reg_dates]
        
        df['일일 판매량'] = (pd.to_numeric(df[t3day], errors='coerce').fillna(0) / divisors).round(0).astype(int)
        v_reorder = pd.to_numeric(df["1차 리오더"], errors='coerce').fillna(0) + pd.to_numeric(df["2차 리오더"], errors='coerce').fillna(0)
        df['권장 발주량'] = ((df['일일 판매량'] * (lead_time + safety_stock)) - (pd.to_numeric(df[avail], errors='coerce') + v_reorder)).clip(lower=0).astype(int)
        
        st.session_state.df_raw = df
        st.success("✅ 분석 완료!")
        st.rerun()

    # 4단계 편집
    st.subheader("📊 4단계: 검색 및 데이터 편집")
    f1, f2 = st.columns([3, 1])
    search = f1.text_input("🔍 상품명 검색")
    filter_mode = f2.selectbox("품절 필터", ["정상만", "품절만", "전체보기"])
    
    df_disp = st.session_state.df_raw.copy()
    if filter_mode == "정상만": df_disp = df_disp[~df_disp[sold_out].astype(str).str.contains('품절', na=False)]
    elif filter_mode == "품절만": df_disp = df_disp[df_disp[sold_out].astype(str).str.contains('품절', na=False)]
    if search: df_disp = df_disp[df_disp[item].astype(str).str.contains(search, na=False)]
    
    edit_cols = [sold_out, vendor, item, option, vendor_item_name, stock, avail, "1차 리오더", "2차 리오더", "일일 판매량", t3day, t1week, '권장 발주량']
    df_final = df_disp[[c for c in edit_cols if c in df_disp.columns]]
    edited_df = st.data_editor(df_final, use_container_width=True, disabled=[c for c in df_final.columns if c not in ["1차 리오더", "2차 리오더"]])
    st.session_state.df_raw.update(edited_df)

    # 5단계: 발주 리스트 요약 및 저장
    st.subheader("📋 5단계: 발주 리스트 요약")
    to_order = st.session_state.df_raw[st.session_state.df_raw['권장 발주량'] > 0].copy()
    
    if not to_order.empty:
        # 화면 출력용
        st.dataframe(to_order, use_container_width=True)
        
        if st.button("💾 기록 저장"):
            date_key = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            # [수정] 명시적으로 데이터프레임 복사본을 저장하여 타입 에러 방지
            st.session_state.history[date_key] = to_order.copy().reset_index(drop=True)
            st.success("✅ 현재 상태가 과거 기록으로 저장되었습니다!")
            st.rerun()

    # 6단계: 과거 데이터 확인
    st.subheader("📜 6단계: 과거 데이터 확인")
    if st.session_state.history:
        # 시간 선택
        s_time = st.selectbox("⏰ 저장된 기록 선택", sorted(st.session_state.history.keys(), reverse=True))
        
        # [수정] 저장된 데이터가 데이터프레임인지 확실히 확인하고 출력
        hist_data = st.session_state.history[s_time]
        if isinstance(hist_data, pd.DataFrame):
            st.dataframe(hist_data, use_container_width=True)
            
            # 엑셀 다운로드 로직
            hist_buf = BytesIO()
            with pd.ExcelWriter(hist_buf, engine='openpyxl') as w: 
                hist_data.to_excel(w, index=False)
            st.download_button("📥 기록 다운로드", data=hist_buf.getvalue(), file_name=f"기록_{s_time}.xlsx")
        else:
            st.error("저장된 데이터 형식이 올바르지 않습니다.")
    else:
        st.info("아직 저장된 과거 기록이 없습니다.")
