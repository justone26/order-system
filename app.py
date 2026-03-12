import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import holidays

st.set_page_config(layout="wide", page_title="재고 관리 시스템")

# 제목
st.markdown('<div style="font-size: 55px; font-weight: 900; color: #000; margin-bottom: 20px;">📦 재고 관리 및 발주 시스템</div>', unsafe_allow_html=True)

if st.button("🔄 시스템 초기화"):
    for key in list(st.session_state.keys()): del st.session_state[key]
    st.rerun()

# [함수] 컬럼 자동 찾기
def get_auto_index(cols, keywords):
    for key in keywords:
        for i, c in enumerate(cols):
            if key in str(c): return i
    return 0

# 데이터 로드
uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'csv'])
if uploaded_file:
    df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
    st.session_state.df_raw = df.loc[:, ~df.columns.duplicated()]
    st.rerun()

if st.session_state.df_raw is not None:
    df = st.session_state.df_raw
    cols = df.columns.tolist()

    # 1단계: 자동 매핑 설정 (항목 복구 완료)
    st.subheader("⚙️ 1단계: 자동 매핑 설정")
    c1, c2 = st.columns(2)
    with c1:
        sold_out = st.selectbox("품절 여부", cols, index=get_auto_index(cols, ['품절', '판매중단']))
        vendor = st.selectbox("공급처", cols, index=get_auto_index(cols, ['공급처', '업체명']))
        item = st.selectbox("상품명", cols, index=get_auto_index(cols, ['상품명', '상품']))
        option = st.selectbox("옵션", cols, index=get_auto_index(cols, ['옵션']))
    with c2:
        avail = st.selectbox("가용재고", cols, index=get_auto_index(cols, ['가용재고', '가용']))
        t3day = st.selectbox("3일 판매합계", cols, index=get_auto_index(cols, ['3일', '최근3일']))
        reorder = st.selectbox("입고예정수량(리오더)", cols, index=get_auto_index(cols, ['리오더', '입고예정']))

    # 2~3단계: 기간 설정 및 분석
    st.subheader("⚙️ 2~3단계: 기간 설정 및 분석")
    l1, l2 = st.columns(2)
    lead_time = l1.number_input("리드타임 (일)", value=0)
    safety_stock = l2.number_input("안전재고 (일)", value=3)
    
    if st.button("🚀 분석 실행", type="primary"):
        # 강제 컬럼 체크 및 생성
        v_avail = pd.to_numeric(df[avail], errors='coerce').fillna(0)
        v_3day = pd.to_numeric(df[t3day], errors='coerce').fillna(0)
        v_reorder = pd.to_numeric(df[reorder], errors='coerce').fillna(0)
        
        # 분석값 계산
        st.session_state.df_raw['권장 발주량'] = ((v_3day / 3) * (lead_time + safety_stock) - (v_avail + v_reorder)).clip(lower=0).round(0)
        st.success("✅ 분석 완료!")
        st.rerun()

    # 4단계: 검색 및 데이터 편집
    st.subheader("📊 4단계: 데이터 편집")
    edit_cols = [sold_out, vendor, item, option, avail, reorder, t3day]
    if '권장 발주량' in df.columns: edit_cols.append('권장 발주량')
    
    # 중복 없이 컬럼 필터링
    df_final = df[list(dict.fromkeys([c for c in edit_cols if c in df.columns]))]
    edited_df = st.data_editor(df_final, use_container_width=True)
    st.session_state.df_raw.update(edited_df)

    # 5단계: 발주 리스트 요약
    st.subheader("📋 5단계: 발주 리스트 요약")
    if '권장 발주량' in df.columns:
        to_order = df[df['권장 발주량'] > 0]
        st.dataframe(to_order, use_container_width=True)

    # 6단계: 과거 기록 확인 (저장 버튼은 5단계 아래)
    st.subheader("📜 6단계: 과거 데이터 확인")
    # ... (기존 기록 코드 그대로 붙여넣기) ...
else:
    st.info("파일을 업로드하면 시스템이 시작됩니다.")
