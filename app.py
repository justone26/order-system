import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 1. 기본 설정
KST = timezone(timedelta(hours=9))
st.set_page_config(layout="wide", page_title="저스트원 재고관리 v1.7")

# 자동 매칭 로직 (사장님 요청 10종)
def get_auto_index(cols, keywords):
    for i, col in enumerate(cols):
        if any(k in str(col) for k in keywords):
            return i
    return 0

# 세션 상태 관리
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'df_final' not in st.session_state: st.session_state.df_final = None
if 'mapping' not in st.session_state: st.session_state.mapping = {}

def reset_all():
    st.session_state.df_raw = None
    st.session_state.df_final = None
    st.session_state.mapping = {}
    st.rerun()

st.title("📦 저스트원 통합 재고 관리 시스템")

tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 관리"])

with tab1:
    # --- [1단계: 데이터 업로드] ---
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

    # --- [2단계: 자동 컬럼 매칭 (5:5)] ---
    if st.session_state.df_raw is not None and st.session_state.df_final is None:
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

        # --- [3단계: 발주 설정 (리드타임 & 안전재고)] ---
        st.divider()
        st.subheader("⚙️ 3단계: 발주 기준 및 분석 설정")
        
        col_lt, col_ss = st.columns(2)
        with col_lt:
            # 사장님 요청: 리드타임 디폴트 7일
            lt_val = st.number_input("⏳ 리드타임 설정 (입고 소요일)", min_value=1, max_value=30, value=7)
        with col_ss:
            # 사장님 요청: 안전재고 디폴트 3일
            ss_val = st.number_input("🛡️ 안전재고 설정 (버퍼 일수)", min_value=0, max_value=60, value=3)

        if st.button("🚀 데이터 분석 시작", use_container_width=True):
            st.session_state.mapping = {
                "sold_out": sold_out, "vendor": vendor, "item": item, "option": option,
                "vendor_item": vendor_item, "reg_date": reg_date, "stock": stock,
                "avail": avail, "t3day": t3day, "t1week": t1week,
                "lt": lt_val, "ss": ss_val
            }
            
            # 실제 분석 로직 실행
            df = st.session_state.df_raw.copy()
            # 1. 일평균 판매량 계산 (7일 발주합계 / 7)
            df['일판매량'] = (pd.to_numeric(df[t1week], errors='coerce').fillna(0) / 7).round(2)
            
            # 2. 필요재고 계산 = 일판매량 * (리드타임 + 안전재고)
            total_days = lt_val + ss_val
            df['필요재고'] = (df['일판매량'] * total_days).round(0).astype(int)
            
            # 3. 권장발주량 = 필요재고 - 가용재고 (0보다 작으면 0으로)
            df['가용재고_num'] = pd.to_numeric(df[avail], errors='coerce').fillna(0)
            df['권장발주량'] = (df['필요재고'] - df['가용재고_num']).clip(lower=0).astype(int)
            
            st.session_state.df_final = df
            st.rerun()

    # --- [4단계: 결과 확인 및 수정] ---
    if st.session_state.df_final is not None:
        m = st.session_state.mapping
        st.divider()
        st.subheader("📝 4단계: 발주 수량 검토")
        st.success(f"✅ 분석 완료 (리드타임 {m['lt']}일 + 안전재고 {m['ss']}일 = 총 {m['lt']+m['ss']}일분 확보 기준)")

        # 사장님 보기 편하게 주요 컬럼만 모아서 에디터 표시
        display_cols = [m['item'], m['option'], m['avail'], m['t1week'], '일판매량', '필요재고', '권장발주량']
        edited_df = st.data_editor(st.session_state.df_final[display_cols], use_container_width=True, hide_index=True)

        if st.button("🗑️ 처음부터 다시 하기", use_container_width=True):
            reset_all()
