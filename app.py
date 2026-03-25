import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone

# 1. 초기 설정 및 시간
KST = timezone(timedelta(hours=9))
now = datetime.now(KST)
today_date = now.strftime("%Y-%m-%d")

st.set_page_config(layout="wide", page_title="저스트원 재고관리 v2.3")

# 리셋 함수
def reset_callback():
    for key in list(st.session_state.keys()):
        del st.session_state[key]

# 자동 매칭 함수
def get_auto_index(cols, keywords):
    for i, col in enumerate(cols):
        if any(k in str(col).strip() for k in keywords):
            return i
    return 0

# 세션 상태 초기화 (핵심: analyzed 추가)
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'df_final' not in st.session_state: st.session_state.df_final = None
if 'mapping' not in st.session_state: st.session_state.mapping = {}
if 'analyzed' not in st.session_state: st.session_state.analyzed = False

st.title("📦 저스트원 통합 재고 관리 시스템")
st.info(f"📅 **분석 기준:** {today_date} (한국 시간)")

tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "🌙 동대문 사입 관리"])

with tab1:
    # --- 1단계: 업로드 ---
    st.subheader("📁 1단계: 데이터 업로드")
    up_file = st.file_uploader("파일 업로드", type=['xlsx', 'xls', 'csv'], key="main_uploader", label_visibility="collapsed")
    
    if st.button("🔄 전체 데이터 초기화", on_click=reset_callback):
        pass

    if up_file is not None and st.session_state.df_raw is None:
        try:
            df = pd.read_csv(up_file) if up_file.name.endswith('.csv') else pd.read_excel(up_file)
            df.columns = df.columns.str.strip()
            st.session_state.df_raw = df
            st.rerun()
        except Exception as e:
            st.error(f"파일 읽기 실패: {e}")

    # --- 2~3단계: 설정 (파일이 있을 때 노출) ---
    if st.session_state.df_raw is not None:
        st.divider()
        st.subheader("🔗 2단계: 컬럼 매칭")
        cols = st.session_state.df_raw.columns.tolist()
        c1, c2 = st.columns(2)
        with c1:
            so = st.selectbox("품절 여부", cols, index=get_auto_index(cols, ['품절']))
            vn = st.selectbox("공급처", cols, index=get_auto_index(cols, ['공급처']))
            it = st.selectbox("상품명", cols, index=get_auto_index(cols, ['상품명']))
            op = st.selectbox("옵션", cols, index=get_auto_index(cols, ['옵션']))
            vi = st.selectbox("공급처 상품명", cols, index=get_auto_index(cols, ['공급처상품명']))
        with c2:
            stk = st.selectbox("정상재고", cols, index=get_auto_index(cols, ['정상재고']))
            av = st.selectbox("가용재고", cols, index=get_auto_index(cols, ['가용재고']))
            t3 = st.selectbox("3일 판매", cols, index=get_auto_index(cols, ['3일']))
            t7 = st.selectbox("7일 판매", cols, index=get_auto_index(cols, ['7일']))

        st.divider()
        st.subheader("⚙️ 3단계: 발주 기준 설정")
        col_lt, col_ss = st.columns(2)
        lt_val = col_lt.number_input("⏳ 리드타임 (일)", min_value=1, value=7)
        ss_val = col_ss.number_input("🛡️ 안전재고 (일)", min_value=0, value=3)

        # [중요] 분석 버튼: 클릭 시 세션의 analyzed를 True로 변경
        if st.button("🚀 데이터 분석 시작", use_container_width=True, type="primary"):
            st.session_state.mapping = {
                'so': so, 'vn': vn, 'it': it, 'op': op, 'vi': vi,
                'st': stk, 'av': av, 't3': t3, 't7': t7,
                'lt': lt_val, 'ss': ss_val
            }
            st.session_state.analyzed = True  # 하단 4~6단계를 여는 스위치
            st.rerun()

    # --- 4~6단계: 결과 (analyzed가 True일 때만 하단에 전개) ---
    if st.session_state.get('analyzed'):
        st.divider()
        st.subheader("📊 4단계: 재고 분석 결과")
        
        m = st.session_state.mapping
        df_work = st.session_state.df_raw.copy()
        
        # 숫자 변환
        for c in [m['st'], m['av'], m['t7'], m['t3']]:
            df_work[c] = pd.to_numeric(df_work[c], errors='coerce').fillna(0).astype(int)
        
        if "리오더 수량" not in df_work.columns: df_work["리오더 수량"] = 0
        
        # 계산 로직
        df_work['일판매량'] = (df_work[m['t7']] / 7).round(1)
        df_work['필요재고'] = (df_work['일판매량'] * (m['lt'] + m['ss'])).round(0).astype(int)
        df_work['권장발주량'] = (df_work['필요재고'] - (df_work[m['av']] + df_work['리오더 수량'])).clip(lower=0).astype(int)
        
        # 결과 표 시각화
        display_cols = [m['it'], m['op'], m['av'], "리오더 수량", "일판매량", "필요재고", "권장발주량"]
        st.data_editor(df_work[display_cols], use_container_width=True, hide_index=True)
        
        st.success("✅ 분석이 완료되었습니다. 위 표에서 수량을 검토하세요.")

        if st.button("🗑️ 분석 결과 숨기기", use_container_width=True):
            st.session_state.analyzed = False
            st.rerun()
