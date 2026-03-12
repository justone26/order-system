import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import holidays

st.set_page_config(layout="wide", page_title="재고 관리 시스템")

# [1] 상태 및 히스토리 초기화
if 'file_key' not in st.session_state: st.session_state.file_key = 0
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'history' not in st.session_state: st.session_state.history = {}

st.title("📦 재고 관리 및 발주 시스템")

if st.button("🔄 시스템 전체 초기화"):
    for key in list(st.session_state.keys()): del st.session_state[key]
    st.session_state.file_key = 1
    st.rerun()

uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'xls', 'csv'], key=f"uploader_{st.session_state.file_key}")

if uploaded_file is not None and st.session_state.df_raw is None:
    try:
        df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
        df = df.loc[:, ~df.columns.duplicated()]
        if "1차 리오더" not in df.columns: df["1차 리오더"] = 0
        if "2차 리오더" not in df.columns: df["2차 리오더"] = 0
        st.session_state.df_raw = df
        st.rerun()
    except Exception as e: st.error(f"파일 오류: {e}")

if st.session_state.df_raw is not None:
    def get_auto_index(cols, keywords):
        for key in keywords:
            for i, c in enumerate(cols):
                if key in str(c): return i
        return 0

    cols = st.session_state.df_raw.columns.tolist()

    # 1단계: 매핑 설정 (공급처 상품명 복구)
    st.subheader("⚙️ 1단계: 자동 매핑 설정")
    c1, c2 = st.columns(2)
    with c1:
        sold_out = st.selectbox("품절 여부", cols, index=get_auto_index(cols, ['품절', '판매중단']))
        vendor = st.selectbox("공급처", cols, index=get_auto_index(cols, ['공급처', '업체명']))
        item = st.selectbox("상품명", cols, index=get_auto_index(cols, ['상품명', '상품']))
        option = st.selectbox("옵션", cols, index=get_auto_index(cols, ['옵션']))
        vendor_item_name = st.selectbox("공급처 상품명", cols, index=get_auto_index(cols, ['공급처상품명', '거래처옵션', '공급처옵션']))
    with c2:
        reg_date_col = st.selectbox("등록일 컬럼", cols, index=get_auto_index(cols, ['등록일', '생성일']))
        stock = st.selectbox("정상재고", cols, index=get_auto_index(cols, ['정상재고', '재고']))
        avail = st.selectbox("가용재고", cols, index=get_auto_index(cols, ['가용재고', '가용']))
        t3day = st.selectbox("3일 발주 합계", cols, index=get_auto_index(cols, ['3일']))
        t1week = st.selectbox("7일 발주 합계", cols, index=get_auto_index(cols, ['7일', '1주']))

    # 2~3단계: 분석 설정
    st.subheader("⚙️ 2~3단계: 분석 설정")
    l1, l2 = st.columns(2)
    lead_time = l1.number_input("리드타임 (일)", value=0)
    safety_stock = l2.number_input("안전재고 (일)", value=3)
    
    if st.button("🚀 분석 실행", type="primary"):
        df = st.session_state.df_raw.copy()
        v_3day = pd.to_numeric(df[t3day], errors='coerce').fillna(0)
        df['일일 판매량'] = (v_3day / 3).round(0).astype(int)
        v_avail = pd.to_numeric(df[avail], errors='coerce').fillna(0)
        v_reorder = pd.to_numeric(df["1차 리오더"], errors='coerce').fillna(0) + pd.to_numeric(df["2차 리오더"], errors='coerce').fillna(0)
        df['권장 발주량'] = ((df['일일 판매량'] * (lead_time + safety_stock)) - (v_avail + v_reorder)).clip(lower=0).astype(int)
        st.session_state.df_raw = df
        st.success("분석 완료!")
        st.rerun()

    # 4단계: 데이터 편집 (공급처 상품명 추가)
    st.subheader("📊 4단계: 데이터 편집")
    edit_cols = [sold_out, vendor, item, option, vendor_item_name, stock, avail, "1차 리오더", "2차 리오더", "일일 판매량", t3day, t1week, '권장 발주량']
    df_final = st.session_state.df_raw[[c for c in edit_cols if c in st.session_state.df_raw.columns]].copy()
    
    # 정수형 변환
    for c in ['일일 판매량', '권장 발주량']:
        if c in df_final.columns: df_final[c] = df_final[c].fillna(0).astype(int)

    edited_df = st.data_editor(
        df_final, 
        use_container_width=True, 
        disabled=[c for c in df_final.columns if c not in ["1차 리오더", "2차 리오더"]]
    )
    st.session_state.df_raw.update(edited_df)

    # 5단계: 발주 리스트 요약 (형식 정렬)
    st.subheader("📋 5단계: 발주 리스트 요약")
    to_order = st.session_state.df_raw[st.session_state.df_raw['권장 발주량'] > 0].copy()
    
    if not to_order.empty:
        display_cols = [vendor, item, option, vendor_item_name, "1차 리오더", "2차 리오더", '권장 발주량']
        to_order_display = to_order[[c for c in display_cols if c in to_order.columns]]
        st.dataframe(to_order_display, use_container_width=True)
        
        c1, c2 = st.columns(2)
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w: to_order_display.to_excel(w, index=False)
        c1.download_button("📥 발주 리스트 다운로드", data=buf.getvalue(), file_name=f"발주서_{datetime.now().strftime('%m%d')}.xlsx")
        
        if c2.button("💾 현재 리스트 기록 저장"):
            time_key = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            st.session_state.history[time_key] = to_order_display.copy()
            st.success(f"{time_key} 기록 저장 완료!")
            st.rerun()

    # 6단계: 과거 데이터 확인
    st.subheader("📜 6단계: 과거 데이터 확인")
    if st.session_state.history:
        s_time = st.selectbox("⏰ 저장된 기록 선택", sorted(st.session_state.history.keys(), reverse=True))
        hist_df = st.session_state.history[s_time]
        if isinstance(hist_df, pd.DataFrame):
            st.dataframe(hist_df, use_container_width=True)
            h_buf = BytesIO()
            with pd.ExcelWriter(h_buf, engine='openpyxl') as w: hist_df.to_excel(w, index=False)
            st.download_button(f"📥 {s_time} 기록 다운로드", data=h_buf.getvalue(), file_name=f"과거기록_{s_time.replace(':', '-')}.xlsx")
    else:
        st.info("저장된 기록이 없습니다.")
