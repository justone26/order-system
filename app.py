import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(layout="wide", page_title="재고 관리 시스템")
st.title("📦 재고 관리 및 발주 시스템")

# [세션 상태에 매핑된 컬럼 정보 보관]
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'mapping' not in st.session_state: 
    st.session_state.mapping = {'품절': None, '공급처': None, '상품명': None, '가용재고': None, '3일합계': None}

uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'csv'])
if uploaded_file is not None and st.session_state.df_raw is None:
    st.session_state.df_raw = pd.read_excel(uploaded_file) if uploaded_file.name.endswith('.xlsx') else pd.read_csv(uploaded_file)
    st.rerun()

if st.session_state.df_raw is not None:
    cols = st.session_state.df_raw.columns.tolist()

    # [1단계: 컬럼 매핑 고정]
    st.subheader("⚙️ 1단계: 컬럼 매핑 (한번 설정 시 유지)")
    c1, c2 = st.columns(2)
    with c1:
        st.session_state.mapping['품절'] = st.selectbox("품절 여부", cols, index=cols.index(st.session_state.mapping['품절']) if st.session_state.mapping['품절'] in cols else 0)
        st.session_state.mapping['공급처'] = st.selectbox("공급처", cols, index=cols.index(st.session_state.mapping['공급처']) if st.session_state.mapping['공급처'] in cols else 0)
        st.session_state.mapping['상품명'] = st.selectbox("상품명", cols, index=cols.index(st.session_state.mapping['상품명']) if st.session_state.mapping['상품명'] in cols else 0)
    with c2:
        st.session_state.mapping['가용재고'] = st.selectbox("가용재고", cols, index=cols.index(st.session_state.mapping['가용재고']) if st.session_state.mapping['가용재고'] in cols else 0)
        st.session_state.mapping['3일합계'] = st.selectbox("3일 발주 합계", cols, index=cols.index(st.session_state.mapping['3일합계']) if st.session_state.mapping['3일합계'] in cols else 0)

    # [2단계: 기간 설정]
    st.subheader("⚙️ 2단계: 기간 설정")
    l1, l2 = st.columns(2)
    lead_time = l1.number_input("리드타임", value=0)
    safety_stock = l2.number_input("안전재고", value=3)

    # [3단계: 분석]
    if st.button("🚀 분석 실행"):
        col_avail = st.session_state.mapping['가용재고']
        col_3day = st.session_state.mapping['3일합계']
        st.session_state.df_raw['권장 발주량'] = ((pd.to_numeric(st.session_state.df_raw[col_3day], errors='coerce') / 3) * (lead_time + safety_stock) - pd.to_numeric(st.session_state.df_raw[col_avail], errors='coerce')).clip(lower=0)
        st.rerun()

    # [4단계: 매핑한 열만 노출]
    st.subheader("📊 4단계: 매핑된 열 편집")
    map_cols = list(st.session_state.mapping.values()) + (['권장 발주량'] if '권장 발주량' in st.session_state.df_raw.columns else [])
    
    f1, f2 = st.columns([2, 1])
    search = f1.text_input("🔍 상품명 검색")
    
    # 선택된 열들만 필터링
    df_disp = st.session_state.df_raw[map_cols].copy()
    if search:
        df_disp = df_disp[df_disp[st.session_state.mapping['상품명']].astype(str).str.contains(search, na=False)]
    
    edited_df = st.data_editor(df_disp, use_container_width=True)
    st.session_state.df_raw.update(edited_df)
