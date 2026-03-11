import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(layout="wide", page_title="재고 관리 시스템")
st.title("📦 재고 관리 및 발주 시스템")

if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'history' not in st.session_state: st.session_state.history = []

def get_idx(cols, keywords):
    for key in keywords:
        for i, c in enumerate(cols):
            if key in str(c): return i
    return 0

# [파일 업로드]
uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'xls', 'csv'])
if uploaded_file is not None and st.session_state.df_raw is None:
    df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
    st.session_state.df_raw = df.loc[:, ~df.columns.duplicated()]
    if "입고예정수량(리오더)" not in st.session_state.df_raw.columns:
        st.session_state.df_raw["입고예정수량(리오더)"] = 0
    st.rerun()

if st.session_state.df_raw is not None:
    cols = st.session_state.df_raw.columns.tolist()

    # [1단계: 매핑]
    st.subheader("⚙️ 1단계: 자동 매핑 설정")
    c1, c2 = st.columns(2)
    with c1:
        sold_out = st.selectbox("품절 여부", cols, index=get_idx(cols, ['품절', '판매중단']))
        vendor = st.selectbox("공급처", cols, index=get_idx(cols, ['공급처', '업체명']))
        item = st.selectbox("상품명", cols, index=get_idx(cols, ['상품명', '상품']))
        option = st.selectbox("옵션", cols, index=get_idx(cols, ['옵션']))
        vendor_opt = st.selectbox("공급처옵션", cols, index=get_idx(cols, ['공급처옵션', '거래처옵션']))
    with c2:
        stock = st.selectbox("정상재고", cols, index=get_idx(cols, ['정상재고', '재고']))
        avail = st.selectbox("가용재고", cols, index=get_idx(cols, ['가용재고', '가용']))
        t3day = st.selectbox("3일 발주 합계", cols, index=get_idx(cols, ['3일', '최근3일']))
        t1week = st.selectbox("1주 발주 합계", cols, index=get_idx(cols, ['1주', '7일', '최근7일']))

    # [2단계: 기간 설정]
    st.subheader("⚙️ 2단계: 기간 설정")
    l1, l2 = st.columns(2)
    lead_time = l1.number_input("리드타임 (일)", value=0)
    safety_stock = l2.number_input("안전재고 (일)", value=3)

    # [3단계: 분석]
    if st.button("🚀 분석 실행", use_container_width=True):
        st.session_state.df_raw['일일 판매량'] = (pd.to_numeric(st.session_state.df_raw[t3day], errors='coerce') / 3).round(0)
        st.session_state.df_raw['권장 발주량'] = (st.session_state.df_raw['일일 판매량'] * (lead_time + safety_stock) - 
                                            (pd.to_numeric(st.session_state.df_raw[avail], errors='coerce') + st.session_state.df_raw["입고예정수량(리오더)"])).clip(lower=0)
        st.rerun()

    # [4단계: 검색 및 데이터 편집]
    st.subheader("📊 4단계: 검색 및 데이터 편집")
    f1, f2 = st.columns([3, 1])
    search = f1.text_input("🔍 상품명 검색")
    filter_mode = f2.selectbox("품절 필터", ["전체보기", "품절만", "정상만"])
    
    # 데이터 필터링 적용
    df_disp = st.session_state.df_raw.copy()
    if filter_mode == "품절만": df_disp = df_disp[df_disp[sold_out].astype(str).str.contains('품절', na=False)]
    elif filter_mode == "정상만": df_disp = df_disp[~df_disp[sold_out].astype(str).str.contains('품절', na=False)]
    if search: df_disp = df_disp[df_disp[item].astype(str).str.contains(search, na=False)]

    # 편집할 컬럼만 추출
    edit_cols = [sold_out, vendor, item, option, vendor_opt, stock, avail, "입고예정수량(리오더)", t3day, t1week, '권장 발주량']
    df_final = df_disp[[c for c in edit_cols if c in df_disp.columns]]
    
    edited_df = st.data_editor(
        df_final, 
        use_container_width=True,
        disabled=[c for c in df_final.columns if c != "입고예정수량(리오더)"]
    )
    st.session_state.df_raw.update(edited_df)

    # [5단계: 기록 및 저장]
    st.subheader("📋 5단계: 발주 리스트 및 기록 저장")
    to_order = st.session_state.df_raw[st.session_state.df_raw['권장 발주량'] > 0]
    st.dataframe(to_order[[item, '권장 발주량', '가용재고']], use_container_width=True)

    if st.button("💾 발주 기록 저장"):
        record = to_order.copy()
        record['저장일시'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        st.session_state.history.append(record)
        st.success("기록 저장 완료!")

    if st.session_state.history:
        for i, hist in enumerate(reversed(st.session_state.history)):
            time_label = hist['저장일시'].iloc[0] if '저장일시' in hist.columns else "기록 데이터"
            with st.expander(f"기록 {len(st.session_state.history)-i} ({time_label})"):
                st.dataframe(hist, use_container_width=True)

    # 다운로드
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        st.session_state.df_raw.to_excel(writer, index=False)
    st.download_button("📥 최종 결과 다운로드", buffer.getvalue(), "결과.xlsx")
