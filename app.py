import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(layout="wide", page_title="재고 관리 시스템")

# [1] 상태 초기화
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'history' not in st.session_state: st.session_state.history = {}

st.title("📦 재고 관리 및 발주 시스템")

# 자동 매핑 함수
def get_auto_index(cols, keywords):
    for key in keywords:
        for i, c in enumerate(cols):
            if key in str(c): return i
    return 0

# [파일 업로드]
uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'xls', 'csv'])
if uploaded_file is not None and st.session_state.df_raw is None:
    df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
    st.session_state.df_raw = df.loc[:, ~df.columns.duplicated()]
    st.rerun()

# [메인 로직]
if st.session_state.df_raw is not None:
    cols = st.session_state.df_raw.columns.tolist()

    # 1단계: 매핑 설정
    st.subheader("⚙️ 1단계: 자동 매핑 설정")
    c1, c2 = st.columns(2)
    sold_out = c1.selectbox("품절 여부", cols, index=get_auto_index(cols, ['품절', '판매중단']))
    vendor = c1.selectbox("공급처", cols, index=get_auto_index(cols, ['공급처', '업체명']))
    item = c1.selectbox("상품명", cols, index=get_auto_index(cols, ['상품명', '상품']))
    option = c1.selectbox("옵션", cols, index=get_auto_index(cols, ['옵션']))
    vendor_item = c1.selectbox("공급처 상품명", cols, index=get_auto_index(cols, ['공급처상품명', '거래처옵션']))
    reg_date = c2.selectbox("등록일", cols, index=get_auto_index(cols, ['등록일', '생성일']))
    stock = c2.selectbox("정상재고", cols, index=get_auto_index(cols, ['정상재고', '재고']))
    avail = c2.selectbox("가용재고", cols, index=get_auto_index(cols, ['가용재고', '가용']))
    t3day = c2.selectbox("3일 발주합계", cols, index=get_auto_index(cols, ['3일']))
    t1week = c2.selectbox("7일 발주합계", cols, index=get_auto_index(cols, ['7일', '1주']))

    # 2~3단계: 분석 설정
    st.subheader("⚙️ 2~3단계: 분석 설정")
    lead_time = st.number_input("리드타임 (일)", value=0)
    safety_stock = st.number_input("안전재고 (일)", value=3)
    
    if st.button("🚀 분석 실행"):
        df = st.session_state.df_raw.copy()
        for col in ["1차 리오더", "2차 리오더"]:
            if col not in df.columns: df[col] = 0
        df['일일 판매량'] = (pd.to_numeric(df[t3day], errors='coerce').fillna(0) / 3).round(0).astype(int)
        df['권장 발주량'] = ((df['일일 판매량'] * (lead_time + safety_stock)) - (pd.to_numeric(df[avail], errors='coerce').fillna(0) + pd.to_numeric(df["1차 리오더"], errors='coerce') + pd.to_numeric(df["2차 리오더"], errors='coerce'))).clip(lower=0).astype(int)
        st.session_state.df_raw = df
        st.rerun()

    # 4단계: 데이터 편집 (필요한 컬럼만 깔끔하게!)
    st.subheader("📊 4단계: 데이터 편집")
    # 보여줄 핵심 컬럼만 정의
    display_cols = [sold_out, vendor, item, option, vendor_item, reg_date, stock, avail, "1차 리오더", "2차 리오더"]
    if '일일 판매량' in st.session_state.df_raw.columns: display_cols += ['일일 판매량', t3day, t1week, '권장 발주량']
    
    # 중복 제거 및 존재하는 컬럼만 필터링
    display_cols = [c for c in dict.fromkeys(display_cols) if c in st.session_state.df_raw.columns]
    
    edited_df = st.data_editor(st.session_state.df_raw[display_cols], use_container_width=True)
    st.session_state.df_raw.update(edited_df)

# 5단계: 발주 리스트 요약 (핵심 정보만 추출)
    st.subheader("📋 5단계: 발주 리스트 요약")
    
    if '권장 발주량' in st.session_state.df_raw.columns:
        # 1. 발주가 필요한 항목만 필터링
        to_order = st.session_state.df_raw[st.session_state.df_raw['권장 발주량'] > 0].copy()
        
        if not to_order.empty:
            # 2. 요약에 필요한 컬럼만 정의
            summary_cols = [vendor, item, option, vendor_item, "권장 발주량"]
            # 리스트에 포함된 컬럼만 골라내기
            summary_cols = [c for c in summary_cols if c in to_order.columns]
            
            # 3. 요약표 출력
            st.dataframe(to_order[summary_cols], use_container_width=True)
            
            c1, c2 = st.columns(2)
            # 4. 엑셀 다운로드 (정확한 요약 데이터만)
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine='openpyxl') as w: 
                to_order[summary_cols].to_excel(w, index=False)
            c1.download_button("📥 요약 발주서 다운로드", data=buf.getvalue(), file_name=f"발주서_{datetime.now().strftime('%m%d').xlsx}")
            
            # 5. 기록 저장
            if c2.button("💾 기록 저장"):
                st.session_state.history[datetime.now().strftime("%Y-%m-%d %H:%M:%S")] = to_order[summary_cols].copy()
                st.success("저장 완료!")
        else:
            st.info("현재 발주할 상품이 없습니다.")
            
    # 6단계: 과거 기록
    st.subheader("📜 6단계: 과거 데이터 확인")
    if st.session_state.history:
        s_time = st.selectbox("⏰ 저장 기록 선택", sorted(st.session_state.history.keys(), reverse=True))
        st.dataframe(st.session_state.history[s_time], use_container_width=True)

