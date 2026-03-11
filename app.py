import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import holidays

st.set_page_config(layout="wide", page_title="재고 관리 시스템")

# [핵심] 제목 크기를 55px로 강제 고정하는 스타일
st.markdown("""
    <style>
    .big-title {
        font-size: 55px !important;
        font-weight: 900 !important;
        color: #000 !important;
        margin: 0 !important;
        padding: 0 !important;
        line-height: 1.2 !important;
    }
    /* 버튼을 제목 크기와 동일하게 투명하게 덮어씌움 */
    div[data-testid="stButton"] button {
        opacity: 0 !important;
        position: absolute !important;
        width: 100% !important;
        height: 80px !important;
        cursor: pointer !important;
    }
    </style>
""", unsafe_allow_html=True)

# 1. 시각적인 제목 출력
st.markdown('<div class="big-title">📦 재고 관리 및 발주 시스템</div>', unsafe_allow_html=True)

# 2. 제목 위치에 투명 버튼을 배치하여 클릭 기능을 수행
if st.button(" "):
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

# 세션 관리
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'history' not in st.session_state: st.session_state.history = {}

def get_auto_index(cols, keywords):
    for key in keywords:
        for i, c in enumerate(cols):
            if key in str(c): return i
    return 0

# 1. 파일 업로드
uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'xls', 'csv'])
if uploaded_file is not None and st.session_state.df_raw is None:
    df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
    st.session_state.df_raw = df.loc[:, ~df.columns.duplicated()]
    if "입고예정수량(리오더)" not in st.session_state.df_raw.columns:
        st.session_state.df_raw["입고예정수량(리오더)"] = 0
    st.rerun()

if st.session_state.df_raw is not None:
    cols = st.session_state.df_raw.columns.tolist()

    # 1단계: 자동 매핑
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

    # 2~3단계: 분석
    st.subheader("⚙️ 2~3단계: 기간 설정 및 분석")
    l1, l2 = st.columns(2)
    lead_time = l1.number_input("리드타임 (일)", value=0)
    safety_stock = l2.number_input("안전재고 (일)", value=3)
    
    if st.button("🚀 분석 실행", type="primary"):
        today = datetime.now()
        kr_holidays = holidays.KR(years=today.year)
        def get_biz_days(start_date):
            if pd.isna(start_date): return 3
            days = pd.date_range(start=start_date, end=today)
            biz_days = [d for d in days if d.weekday() < 5 and d not in kr_holidays]
            return max(1, len(biz_days))
        
        reg_dates = pd.to_datetime(st.session_state.df_raw[reg_date_col], errors='coerce')
        divisors = [min(3, get_biz_days(rd)) for rd in reg_dates]
        v_avail = pd.to_numeric(st.session_state.df_raw[avail], errors='coerce').fillna(0)
        v_3day = pd.to_numeric(st.session_state.df_raw[t3day], errors='coerce').fillna(0)
        v_reorder = pd.to_numeric(st.session_state.df_raw["입고예정수량(리오더)"], errors='coerce').fillna(0)
        
        st.session_state.df_raw['일일 판매량'] = (v_3day / divisors).round(1)
        st.session_state.df_raw['권장 발주량'] = ((st.session_state.df_raw['일일 판매량'] * (lead_time + safety_stock)) - (v_avail + v_reorder)).clip(lower=0).round(0)
        st.success("✅ 분석 완료!")
        st.rerun()

    # 4단계: 검색 및 데이터 편집
    st.subheader("📊 4단계: 검색 및 데이터 편집")
    f1, f2 = st.columns([3, 1])
    search = f1.text_input("🔍 상품명 검색")
    filter_mode = f2.selectbox("품절 필터", ["정상만", "품절만", "전체보기"])
    
    df_disp = st.session_state.df_raw.copy()
    if filter_mode == "정상만": df_disp = df_disp[~df_disp[sold_out].astype(str).str.contains('품절', na=False)]
    elif filter_mode == "품절만": df_disp = df_disp[df_disp[sold_out].astype(str).str.contains('품절', na=False)]
    if search: df_disp = df_disp[df_disp[item].astype(str).str.contains(search, na=False)]
    
    edit_cols = [sold_out, vendor, item, option, vendor_item_name, stock, avail, "입고예정수량(리오더)", t3day, t1week, '권장 발주량']
    df_final = df_disp[[c for c in edit_cols if c in df_disp.columns]]
    edited_df = st.data_editor(df_final, use_container_width=True, disabled=[c for c in df_final.columns if c != "입고예정수량(리오더)"])
    st.session_state.df_raw.update(edited_df)

    # 5단계: 요약
    st.subheader("📋 5단계: 발주 리스트 요약")
    if '권장 발주량' in st.session_state.df_raw.columns:
        to_order = st.session_state.df_raw[st.session_state.df_raw['권장 발주량'] > 0].copy()
        if not to_order.empty:
            display_df = to_order[[vendor, item, option, vendor_item_name, '권장 발주량']].rename(columns={
                vendor: "공급처", item: "상품명", option: "옵션", vendor_item_name: "공급처 상품명", '권장 발주량': "최종 발주량"
            })
            edited_order = st.data_editor(display_df, use_container_width=True)
            c1, c2 = st.columns(2)
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine='openpyxl') as w: edited_order.to_excel(w, index=False)
            c1.download_button("📥 발주 리스트 다운로드", data=buf.getvalue(), file_name="발주서.xlsx")
            if c2.button("💾 기록 저장"):
                date_key = datetime.now().strftime("%Y-%m-%d")
                record = to_order.copy()
                record['저장시각'] = datetime.now().strftime("%H:%M:%S")
                if date_key not in st.session_state.history: st.session_state.history[date_key] = []
                st.session_state.history[date_key].append(record)
                st.success("기록 저장 완료!")

    # 6단계: 과거 데이터 확인
    st.subheader("📜 6단계: 과거 데이터 확인")
    if st.session_state.history:
        s_date = st.selectbox("📅 날짜 선택", sorted(st.session_state.history.keys(), reverse=True))
        s_time = st.selectbox("⏰ 시간 선택", [h['저장시각'].iloc[0] for h in st.session_state.history[s_date]][::-1])
        for h in st.session_state.history[s_date]:
            if h['저장시각'].iloc[0] == s_time:
                st.dataframe(h[[c for c in edit_cols if c in h.columns]], use_container_width=True)
                hist_buf = BytesIO()
                with pd.ExcelWriter(hist_buf, engine='openpyxl') as w: h.to_excel(w, index=False)
                st.download_button("📥 기록 다운로드", data=hist_buf.getvalue(), file_name=f"기록_{s_date}_{s_time}.xlsx")

    st.divider()
    all_buf = BytesIO()
    with pd.ExcelWriter(all_buf, engine='openpyxl') as w: st.session_state.df_raw.to_excel(w, index=False)
    st.download_button("📥 전체 원본 데이터 다운로드", data=all_buf.getvalue(), file_name="최종결과데이터.xlsx")
