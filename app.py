import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(layout="wide", page_title="재고 관리 시스템")
st.title("📦 재고 관리 및 발주 시스템")

if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'history' not in st.session_state: st.session_state.history = {}

def get_idx(cols, keywords):
    for key in keywords:
        for i, c in enumerate(cols):
            if key in str(c): return i
    return 0

# 1. 업로드
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
    sold_out = c1.selectbox("품절 여부", cols, index=get_idx(cols, ['품절', '판매중단']))
    vendor = c1.selectbox("공급처", cols, index=get_idx(cols, ['공급처', '업체명']))
    item = c1.selectbox("상품명", cols, index=get_idx(cols, ['상품명', '상품']))
    option = c1.selectbox("옵션", cols, index=get_idx(cols, ['옵션']))
    vendor_opt = c1.selectbox("공급처옵션", cols, index=get_idx(cols, ['공급처옵션', '거래처옵션']))
    stock = c2.selectbox("정상재고", cols, index=get_idx(cols, ['정상재고', '재고']))
    avail = c2.selectbox("가용재고", cols, index=get_idx(cols, ['가용재고', '가용']))
    t3day = c2.selectbox("3일 발주 합계", cols, index=get_idx(cols, ['3일', '최근3일']))
    t1week = c2.selectbox("1주 발주 합계", cols, index=get_idx(cols, ['1주', '7일', '최근7일']))

    # [2단계: 파라미터]
    st.subheader("⚙️ 2단계: 파라미터 설정")
    l1, l2 = st.columns(2)
    lead_time = l1.number_input("리드타임", value=0)
    safety_stock = l2.number_input("안전재고", value=3)

    # [3단계: 분석]
    st.subheader("⚙️ 3단계: 분석 실행")
    if st.button("🚀 분석 실행"):
        st.session_state.df_raw['일일 판매량'] = (pd.to_numeric(st.session_state.df_raw[t3day], errors='coerce') / 3).round(0)
        st.session_state.df_raw['권장 발주량'] = (st.session_state.df_raw['일일 판매량'] * (lead_time + safety_stock) - 
                                            (pd.to_numeric(st.session_state.df_raw[avail], errors='coerce') + st.session_state.df_raw["입고예정수량(리오더)"])).clip(lower=0)
        st.rerun()

    # [4단계: 데이터 편집 (검색/필터 복구)]
    st.subheader("📊 4단계: 검색 및 데이터 편집")
    f1, f2 = st.columns([3, 1])
    search = f1.text_input("🔍 상품명 검색")
    filter_mode = f2.selectbox("품절 필터", ["전체보기", "품절만", "정상만"])
    
    df_disp = st.session_state.df_raw.copy()
    if filter_mode == "품절만": df_disp = df_disp[df_disp[sold_out].astype(str).str.contains('품절', na=False)]
    elif filter_mode == "정상만": df_disp = df_disp[~df_disp[sold_out].astype(str).str.contains('품절', na=False)]
    if search: df_disp = df_disp[df_disp[item].astype(str).str.contains(search, na=False)]

    edit_cols = [sold_out, vendor, item, option, vendor_opt, stock, avail, "입고예정수량(리오더)", t3day, t1week, '권장 발주량']
    df_final = df_disp[[c for c in edit_cols if c in df_disp.columns]]
    edited_df = st.data_editor(df_final, use_container_width=True, disabled=[c for c in df_final.columns if c != "입고예정수량(리오더)"])
    st.session_state.df_raw.update(edited_df)

    # [5단계: 발주 요약]
    st.subheader("📋 5단계: 발주 필요 리스트 요약")
    if '권장 발주량' in st.session_state.df_raw.columns:
        to_order = st.session_state.df_raw[st.session_state.df_raw['권장 발주량'] > 0]
        st.dataframe(to_order[edit_cols], use_container_width=True)
        if st.button("💾 리스트 기록 저장"):
            date_key = datetime.now().strftime("%Y-%m-%d")
            record = to_order[edit_cols].copy()
            record['저장시각'] = datetime.now().strftime("%H:%M:%S")
            if date_key not in st.session_state.history: st.session_state.history[date_key] = []
            st.session_state.history[date_key].append(record)
            st.success("저장 완료!")
    else: st.warning("3단계 분석을 먼저 실행해주세요.")

    # [6단계: 과거 확인]
    st.subheader("📜 6단계: 과거 데이터 확인")
    if st.session_state.history:
        selected_date = st.selectbox("조회할 날짜 선택", sorted(st.session_state.history.keys(), reverse=True))
        for hist in st.session_state.history[selected_date]:
            st.dataframe(hist, use_container_width=True)

    st.download_button("📥 최종 데이터 다운로드", data=BytesIO(st.session_state.df_raw.to_excel(index=False).encode('utf-8')), file_name="결과.xlsx")


# --- [🌙 탭 2: 동대문 사입 관리] ---
with tab2:
    st.subheader("🌙 동대문 사입 및 미납 관리")
    dong_file = st.file_uploader("동대문 주문 리스트 업로드", type=['xlsx', 'csv'], key="dong_tab_upload")
    if dong_file:
        if "last_file_name" not in st.session_state or st.session_state.last_file_name != dong_file.name:
            df = pd.read_excel(dong_file)
            df.columns = df.columns.str.strip()
            required_cols = ['선택', '품절', '상품명', '공급처', '공급처상품명', '정상재고', '가용재고', '판매수량', '발주수량', '가중율', '3일판매']
            for col in required_cols:
                if col not in df.columns: df[col] = 0 if col not in ['선택', '품절', '상품명', '공급처', '공급처상품명'] else ""
            for col in ['정상재고', '가용재고', '3일판매']: df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            df['판매수량'] = (df['정상재고'] - df['가용재고']).clip(lower=0)
            df['가중율'] = df['판매수량'].apply(lambda n: 2.0 if n >= 10 else (1.5 if n >= 6 else (1.2 if n >= 3 else 1.0)))
            df['발주수량'] = (df['판매수량'] * df['가중율']).astype(int)
            st.session_state.df_dong_current = df[required_cols]
            st.session_state.last_file_name = dong_file.name

        df_display = st.session_state.df_dong_current.copy()
        search_query = st.text_input("상품명 검색 (사입)")
        if search_query: df_display = df_display[df_display['상품명'].astype(str).str.contains(search_query, case=False, na=False)]
        
        df_display['선택'] = df_display['선택'].astype(bool)
        edited_df = st.data_editor(df_display, use_container_width=True, key="dong_editor")
        
        st.divider()
        c1, c2, c3 = st.columns(3)
        add_val = c1.number_input("추가 수량", value=1, min_value=1)
        if c2.button("🚀 선택 상품 수량 더하기"):
            selected = edited_df[edited_df['선택'] == True].index
            for idx in selected: st.session_state.df_dong_current.at[idx, '발주수량'] += add_val
            st.rerun()
        csv = edited_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
        c3.download_button("📥 엑셀 다운로드", csv, "사입리스트.csv", "text/csv")
