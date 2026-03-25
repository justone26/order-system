import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 1. 기본 설정
KST = timezone(timedelta(hours=9))
st.set_page_config(layout="wide", page_title="저스트원 재고관리 v1.5")

# 사장님이 사용하시는 자동 매칭 로직 (get_auto_index)
def get_auto_index(cols, keywords):
    for i, col in enumerate(cols):
        if any(k in str(col) for k in keywords):
            return i
    return 0

# 세션 상태 관리
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'df_final' not in st.session_state: st.session_state.df_final = None

def reset_all():
    st.session_state.df_raw = None
    st.session_state.df_final = None
    st.rerun()

st.title("📦 저스트원 통합 재고 관리 시스템")

tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 관리"])

with tab1:
    # --- [1단계: 데이터 업로드 & 초기화] ---
    st.subheader("📁 1단계: 데이터 업로드")
    up_file = st.file_uploader("파일을 선택하세요", type=['xlsx', 'xls', 'csv'], key="up_key", label_visibility="collapsed")
    
    col_btn, _ = st.columns([1, 3])
    with col_btn:
        if st.button("🔄 전체 데이터 초기화", use_container_width=True):
            reset_all()

    if up_file and st.session_state.df_raw is None:
        try:
            df = pd.read_csv(up_file) if up_file.name.endswith('.csv') else pd.read_excel(up_file)
            df.columns = df.columns.str.strip()
            st.session_state.df_raw = df
            st.rerun()
        except Exception as e:
            st.error(f"파일 읽기 실패: {e}")

    # --- [2단계: 사장님 요청 10가지 항목 (5:5 배치)] ---
    if st.session_state.df_raw is not None and st.session_state.df_final is None:
        st.divider()
        st.subheader("🔗 2단계: 자동 컬럼 매칭 (좌우 5:5 정렬)")
        cols = st.session_state.df_raw.columns.tolist()
        
        c1, c2 = st.columns(2)
        
        with c1:
            st.info("📍 기본 정보 (C1)")
            sold_out = st.selectbox("품절 여부", cols, index=get_auto_index(cols, ['품절', '판매중단']))
            vendor = st.selectbox("공급처", cols, index=get_auto_index(cols, ['공급처', '업체명']))
            item = st.selectbox("상품명", cols, index=get_auto_index(cols, ['상품명', '상품']))
            option = st.selectbox("옵션", cols, index=get_auto_index(cols, ['옵션']))
            vendor_item = st.selectbox("공급처 상품명", cols, index=get_auto_index(cols, ['공급처상품명', '거래처옵션']))
            
        with c2:
            st.info("📊 재고/판매 정보 (C2)")
            reg_date = st.selectbox("등록일", cols, index=get_auto_index(cols, ['등록일', '생성일']))
            stock = st.selectbox("정상재고", cols, index=get_auto_index(cols, ['정상재고', '재고']))
            avail = st.selectbox("가용재고", cols, index=get_auto_index(cols, ['가용재고', '가용']))
            t3day = st.selectbox("3일 발주합계", cols, index=get_auto_index(cols, ['3일']))
            t1week = st.selectbox("7일 발주합계", cols, index=get_auto_index(cols, ['7일', '1주']))

        # --- [3단계: 분석 실행] ---
        st.divider()
        st.subheader("⚙️ 3단계: 분석 및 계산 실행")
        if st.button("🚀 분석 시작 (수량 동기화)", use_container_width=True):
            # 매칭 정보를 세션에 저장 (발주량 계산용)
            st.session_state.mapping = {
                "sold_out": sold_out, "vendor": vendor, "item": item, "option": option,
                "vendor_item": vendor_item, "reg_date": reg_date, "stock": stock,
                "avail": avail, "t3day": t3day, "t1week": t1week
            }
            st.session_state.df_final = st.session_state.df_raw.copy() 
            st.rerun()

    # --- [4~6단계: 수정 및 저장] ---
    if st.session_state.df_final is not None:
        st.divider()
        st.success("✅ 매칭 완료! 수량을 확인하고 수정하세요.")
        # 여기에 편집기 및 구글시트 업데이트 로직 추가
