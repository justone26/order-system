import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 1. 기본 설정
KST = timezone(timedelta(hours=9))
st.set_page_config(layout="wide", page_title="저스트원 재고관리 v1.8")

# 자동 매칭 로직 (사장님 요청 10종)
def get_auto_index(cols, keywords):
    for i, col in enumerate(cols):
        if any(k in str(col).strip() for k in keywords):
            return i
    return 0

# [수정] 먹통 방지용 강력 초기화 함수
def reset_all():
    # 세션에 저장된 모든 데이터를 삭제
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    # 즉시 화면 갱신
    st.rerun()

# 세션 상태 초기화 (최초 1회)
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'df_final' not in st.session_state: st.session_state.df_final = None

st.title("📦 저스트원 통합 재고 관리 시스템")

tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 관리"])

with tab1:
    # --- [1단계: 데이터 업로드] ---
    st.subheader("📁 1단계: 데이터 업로드")
    
    # 파일 업로더 (key를 고정해서 초기화 시 같이 날아가게 함)
    up_file = st.file_uploader("파일을 선택하세요", type=['xlsx', 'xls', 'csv'], key="file_uploader_key", label_visibility="collapsed")
    
    # [사장님 요청 위치] 업로드 칸 바로 아래 왼쪽
    col_btn, _ = st.columns([1, 3])
    with col_btn:
        if st.button("🔄 전체 데이터 초기화", use_container_width=True, on_click=reset_all):
            # on_click에 함수를 직접 연결해서 더 확실하게 작동하게 함
            pass

    # 파일 읽기 (업로드된 파일이 있고, 아직 데이터프레임이 생성 안 됐을 때만)
    if up_file is not None and st.session_state.df_raw is None:
        try:
            df = pd.read_csv(up_file) if up_file.name.endswith('.csv') else pd.read_excel(up_file)
            df.columns = df.columns.str.strip()
            st.session_state.df_raw = df
            st.rerun()
        except Exception as e:
            st.error(f"파일 읽기 실패: {e}")

    # --- [2~3단계: 매칭 및 설정] ---
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

        st.divider()
        st.subheader("⚙️ 3단계: 발주 기준 설정")
        col_lt, col_ss = st.columns(2)
        with col_lt:
            lt_val = st.number_input("⏳ 리드타임 설정 (7일 디폴트)", min_value=1, value=7)
        with col_ss:
            ss_val = st.number_input("🛡️ 안전재고 설정 (3일 디폴트)", min_value=0, value=3)

        if st.button("🚀 분석 시작", use_container_width=True):
            st.session_state.mapping = {
                "sold_out": sold_out, "vendor": vendor, "item": item, "option": option,
                "vendor_item": vendor_item, "reg_date": reg_date, "stock": stock,
                "avail": avail, "t3day": t3day, "t1week": t1week, "lt": lt_val, "ss": ss_val
            }
            # 계산 엔진
            df = st.session_state.df_raw.copy()
            df['일판매량'] = (pd.to_numeric(df[t1week], errors='coerce').fillna(0) / 7).round(2)
            df['필요재고'] = (df['일판매량'] * (lt_val + ss_val)).round(0).astype(int)
            df['가용재고_num'] = pd.to_numeric(df[avail], errors='coerce').fillna(0)
            df['권장발주량'] = (df['필요재고'] - df['가용재고_num']).clip(lower=0).astype(int)
            
            st.session_state.df_final = df
            st.rerun()

    # --- [4단계: 결과 확인] ---
    if st.session_state.df_final is not None:
        st.divider()
        st.success("✅ 분석 완료!")
        m = st.session_state.mapping
        display_cols = [m['item'], m['option'], m['avail'], m['t1week'], '일판매량', '필요재고', '권장발주량']
        st.data_editor(st.session_state.df_final[display_cols], use_container_width=True, hide_index=True)

        if st.button("🗑️ 처음부터 다시 하기", on_click=reset_all, use_container_width=True):
            pass
