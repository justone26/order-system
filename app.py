import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(layout="wide", page_title="재고 관리 시스템")

# [핵심] 시스템 전체 초기화 (업로드 파일까지 제거)
if st.button("🔄 시스템 전체 초기화"):
    # 세션 상태 전부 삭제
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    
    # 1. 파일 업로더 초기화를 위해 쿼리 파라미터 활용
    # 이렇게 하면 페이지를 새로 고치면서 업로더도 깨끗해짐
    st.rerun()

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

    # 4단계: 데이터 편집
    st.subheader("📊 4단계: 데이터 편집")
    f1, f2 = st.columns([3, 1])
    search_query = f1.text_input("🔍 상품명 검색")
    filter_mode = f2.selectbox("품절 필터", ["정상만", "품절만", "전체보기"], index=0)
    
    df_working = st.session_state.df_raw.copy()
    if filter_mode == "정상만": df_working = df_working[~df_working[sold_out].astype(str).str.contains('품절', na=False)]
    elif filter_mode == "품절만": df_working = df_working[df_working[sold_out].astype(str).str.contains('품절', na=False)]
    if search_query: df_working = df_working[df_working[item].astype(str).str.contains(search_query, na=False)]
    
    display_cols = [sold_out, vendor, item, option, vendor_item, reg_date, stock, avail, "1차 리오더", "2차 리오더"]
    if '일일 판매량' in df_working.columns: display_cols += ['일일 판매량', t3day, t1week, '권장 발주량']
    display_cols = [c for c in dict.fromkeys(display_cols) if c in df_working.columns]
    
    edited_df = st.data_editor(df_working[display_cols], use_container_width=False, key="main_editor")
    st.session_state.df_raw.update(edited_df)

    # 5단계: 발주 리스트 요약
    st.subheader("📋 5단계: 발주 리스트 요약")
    if '권장 발주량' in st.session_state.df_raw.columns:
        to_order = st.session_state.df_raw[st.session_state.df_raw['권장 발주량'] > 0].copy()
        if not to_order.empty:
            st.warning(f"🚨 알림: 발주가 필요한 상품이 {len(to_order)}개 있습니다!")
            summary_cols = [vendor, item, option, vendor_item, "권장 발주량"]
            st.dataframe(to_order[summary_cols], use_container_width=False)
            c1, c2 = st.columns(2)
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine='openpyxl') as w: to_order[summary_cols].to_excel(w, index=False)
            c1.download_button("📥 요약 발주서 다운로드", data=buf.getvalue(), file_name=f"발주서_{datetime.now().strftime('%m%d')}.xlsx")
            if c2.button("💾 기록 저장"):
                st.session_state.history[datetime.now().strftime("%Y-%m-%d %H:%M:%S")] = to_order[summary_cols].copy()
                st.success("저장 완료!")
        else: st.success("✅ 현재 모든 상품의 재고가 충분합니다.")

    # 6단계: 과거 기록
    st.subheader("📜 6단계: 과거 데이터 확인")
    if st.session_state.history:
        s_time = st.selectbox("⏰ 저장 기록 선택", sorted(st.session_state.history.keys(), reverse=True))
        st.dataframe(st.session_state.history[s_time], use_container_width=False)

