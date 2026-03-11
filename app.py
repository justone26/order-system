import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# 페이지 설정
st.set_page_config(layout="wide", page_title="재고 관리 시스템")
st.title("📦 재고 관리 및 발주 시스템")

# 세션 관리
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'history' not in st.session_state: st.session_state.history = {}

def get_idx(cols, keywords):
    for key in keywords:
        for i, c in enumerate(cols):
            if key in str(c): return i
    return 0

# 1. 파일 업로드
uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'xls', 'csv'])
if uploaded_file is not None and st.session_state.df_raw is None:
    df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
    st.session_state.df_raw = df.loc[:, ~df.columns.duplicated()]
    if "입고예정수량(리오더)" not in st.session_state.df_raw.columns:
        st.session_state.df_raw["입고예정수량(리오더)"] = 0
    st.rerun()

if st.session_state.df_raw is not None:
    cols = st.session_state.df_raw.columns.tolist()

    # 1단계: 매핑 설정
    st.subheader("⚙️ 1단계: 자동 매핑 설정")
    c1, c2 = st.columns(2)
    with c1:
        sold_out = st.selectbox("품절 여부", cols, index=get_idx(cols, ['품절', '판매중단']))
        vendor = st.selectbox("공급처", cols, index=get_idx(cols, ['공급처', '업체명']))
        item = st.selectbox("상품명", cols, index=get_idx(cols, ['상품명', '상품']))
        option = st.selectbox("옵션", cols, index=get_idx(cols, ['옵션']))
        vendor_item_name = st.selectbox("공급처 상품명", cols, index=get_idx(cols, ['공급처상품명', '거래처옵션', '공급처옵션']))
    with c2:
        stock = st.selectbox("정상재고", cols, index=get_idx(cols, ['정상재고', '재고']))
        avail = st.selectbox("가용재고", cols, index=get_idx(cols, ['가용재고', '가용']))
        t3day = st.selectbox("3일 발주 합계", cols, index=get_idx(cols, ['3일', '최근3일']))
        t1week = st.selectbox("1주 발주 합계", cols, index=get_idx(cols, ['1주', '7일', '최근7일']))

    # 2~3단계: 분석 실행
    st.subheader("⚙️ 2~3단계: 기간 설정 및 분석")
    l1, l2 = st.columns(2)
    lead_time = l1.number_input("리드타임 (일)", value=0, key="lt_input")
    safety_stock = l2.number_input("안전재고 (일)", value=3, key="ss_input")
    
    if st.button("🚀 분석 실행", type="primary"):
        v_avail = pd.to_numeric(st.session_state.df_raw[avail], errors='coerce').fillna(0)
        v_3day = pd.to_numeric(st.session_state.df_raw[t3day], errors='coerce').fillna(0)
        v_reorder = pd.to_numeric(st.session_state.df_raw["입고예정수량(리오더)"], errors='coerce').fillna(0)
        
        st.session_state.df_raw['일일 판매량'] = (v_3day / 3).round(1)
        calc_val = (st.session_state.df_raw['일일 판매량'] * (lead_time + safety_stock)) - (v_avail + v_reorder)
        st.session_state.df_raw['권장 발주량'] = calc_val.clip(lower=0).round(0)
        
        st.success("분석이 완료되었습니다!")
        st.rerun()

    # 4단계: 데이터 편집
    st.subheader("📊 4단계: 검색 및 데이터 편집")
    f1, f2 = st.columns([3, 1])
    search = f1.text_input("🔍 상품명 검색")
    filter_mode = f2.selectbox("품절 필터", ["전체보기", "품절만", "정상만"])
    
    df_disp = st.session_state.df_raw.copy()
    if filter_mode == "품절만": df_disp = df_disp[df_disp[sold_out].astype(str).str.contains('품절', na=False)]
    elif filter_mode == "정상만": df_disp = df_disp[~df_disp[sold_out].astype(str).str.contains('품절', na=False)]
    if search: df_disp = df_disp[df_disp[item].astype(str).str.contains(search, na=False)]
    
    edit_cols = [sold_out, vendor, item, option, vendor_item_name, stock, avail, "입고예정수량(리오더)", t3day, t1week, '권장 발주량']
    df_final = df_disp[[c for c in edit_cols if c in df_disp.columns]]
    
    edited_df = st.data_editor(df_final, use_container_width=True, key="main_editor",
                               disabled=[c for c in df_final.columns if c != "입고예정수량(리오더)"])
    
    if not edited_df.equals(df_final):
        st.session_state.df_raw.update(edited_df)

    # 5단계: 발주 요약 및 저장
    st.subheader("📋 5단계: 발주 필요 리스트 요약")
    if '권장 발주량' in st.session_state.df_raw.columns:
        to_order = st.session_state.df_raw[st.session_state.df_raw['권장 발주량'] > 0].copy()
        
        if not to_order.empty:
            display_df = to_order[[vendor, item, option, vendor_item_name, '권장 발주량']].rename(columns={
                vendor: "공급처", item: "상품명", option: "옵션", vendor_item_name: "공급처 상품명", '권장 발주량': "최종 발주량"
            })
            
            edited_order = st.data_editor(display_df, use_container_width=True, key="order_editor")
            
            c1, c2 = st.columns(2)
            order_buf = BytesIO()
            with pd.ExcelWriter(order_buf, engine='openpyxl') as writer:
                edited_order.to_excel(writer, index=False)
            c1.download_button("📥 발주 리스트 다운로드", data=order_buf.getvalue(), file_name=f"발주서_{datetime.now().strftime('%m%d')}.xlsx")
            
            if c2.button("💾 이 리스트 기록 저장"):
                date_key = datetime.now().strftime("%Y-%m-%d")
                record = to_order.copy()
                record['저장시각'] = datetime.now().strftime("%H:%M:%S")
                if date_key not in st.session_state.history: st.session_state.history[date_key] = []
                st.session_state.history[date_key].append(record)
                st.success("기록 저장 완료!")
        else:
            st.info("발주 권장 항목이 없습니다.")

    # 6단계: 과거 확인 (스크롤 최적화 버전)
    st.subheader("📜 6단계: 과거 데이터 확인")
    if st.session_state.history:
        # 1. 날짜 먼저 선택
        h_dates = sorted(st.session_state.history.keys(), reverse=True)
        s_date = st.selectbox("📅 조회할 날짜 선택", h_dates)
        
        # 2. 해당 날짜의 저장 시각들을 리스트로 추출
        time_options = [h['저장시각'].iloc[0] for h in st.session_state.history[s_date]]
        # 인덱스를 역순으로 하여 최근 저장 시각이 먼저 나오게 함
        selected_time = st.selectbox("⏰ 저장 시각 선택", time_options[::-1])
        
        # 3. 선택한 시각의 데이터만 찾아서 표시
        for hist in st.session_state.history[s_date]:
            if hist['저장시각'].iloc[0] == selected_time:
                f_cols = [c for c in edit_cols if c in hist.columns]
                st.write(f"✅ **{s_date} {selected_time}** 저장 기록")
                st.dataframe(hist[f_cols], use_container_width=True)
                
                # 해당 시점 데이터만 별도 다운로드 기능
                hist_buf = BytesIO()
                with pd.ExcelWriter(hist_buf, engine='openpyxl') as writer:
                    hist[f_cols].to_excel(writer, index=False)
                st.download_button(f"📥 {selected_time} 기록 다운로드", data=hist_buf.getvalue(), 
                                   file_name=f"발주기록_{s_date}_{selected_time}.xlsx", key=f"dl_{selected_time}")
                break
    else:
        st.info("저장된 과거 기록이 없습니다.")

    # 전체 데이터 다운로드
    st.divider()
    all_buf = BytesIO()
    with pd.ExcelWriter(all_buf, engine='openpyxl') as writer:
        st.session_state.df_raw.to_excel(writer, index=False)
    st.download_button("📥 전체 원본 데이터 다운로드", data=all_buf.getvalue(), file_name="최종결과데이터.xlsx")
