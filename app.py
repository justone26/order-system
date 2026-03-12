import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import holidays

st.set_page_config(layout="wide", page_title="재고 관리 시스템")

# [디자인] 제목 고정
st.markdown("""
    <div style="font-size: 55px; font-weight: 900; color: #000; margin-bottom: 20px;">
        📦 재고 관리 및 발주 시스템
    </div>
""", unsafe_allow_html=True)

if st.button("🔄 시스템 초기화"):
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

st.divider()

# 세션 관리 및 데이터 병합 함수
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'history' not in st.session_state: st.session_state.history = {}

def merge_new_data(old_df, new_df):
    if old_df is None: return new_df
    old_df['key'] = old_df['상품명'].astype(str) + "_" + old_df['옵션'].astype(str)
    new_df['key'] = new_df['상품명'].astype(str) + "_" + new_df['옵션'].astype(str)
    manual_data = old_df[['key', '1차 입고예정', '2차 입고예정']]
    merged_df = pd.merge(new_df, manual_data, on='key', how='left')
    merged_df['1차 입고예정'] = merged_df['1차 입고예정'].fillna(0)
    merged_df['2차 입고예정'] = merged_df['2차 입고예정'].fillna(0)
    return merged_df.drop(columns=['key'])

# [핵심] 가용재고 증가 시 입고예정 자동 차감 로직
def auto_reduce_reorder(old_df, edited_df):
    # 가용재고 변경분 계산
    diff = edited_df['가용재고'] - old_df['가용재고']
    
    for i in diff[diff > 0].index:
        amount_to_reduce = diff[i]
        # 1차에서 먼저 차감
        reduce_1st = min(edited_df.at[i, '1차 입고예정'], amount_to_reduce)
        edited_df.at[i, '1차 입고예정'] -= reduce_1st
        amount_to_reduce -= reduce_1st
        # 남은 양 2차에서 차감
        if amount_to_reduce > 0:
            reduce_2nd = min(edited_df.at[i, '2차 입고예정'], amount_to_reduce)
            edited_df.at[i, '2차 입고예정'] -= reduce_2nd
    return edited_df

# 1. 파일 업로드
uploaded_file = st.file_uploader("엑셀/CSV 업로드", type=['xlsx', 'xls', 'csv'])
if uploaded_file is not None:
    new_df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
    new_df = new_df.loc[:, ~new_df.columns.duplicated()]
    # 기존 수기 데이터 보존하며 병합
    st.session_state.df_raw = merge_new_data(st.session_state.df_raw, new_df)
    st.rerun()

if st.session_state.df_raw is not None:
    cols = st.session_state.df_raw.columns.tolist()
    
    # [매핑 생략 - 기존 로직 유지]
    # ... (매핑 설정 단계) ...

    # 4단계: 데이터 편집 (자동 차감 적용)
    st.subheader("📊 4단계: 검색 및 데이터 편집")
    edit_cols = ['품절여부', '공급처', '상품명', '옵션', '가용재고', '1차 입고예정', '2차 입고예정', '권장 발주량']
    # 필요한 컬럼만 추출
    df_final = st.session_state.df_raw[[c for c in edit_cols if c in st.session_state.df_raw.columns]]
    
    # 편집기 실행
    edited_df = st.data_editor(df_final, use_container_width=True)
    
    # 변경 사항이 있으면 자동 차감 로직 실행 후 세션 업데이트
    if not edited_df.equals(df_final):
        updated_df = auto_reduce_reorder(df_final, edited_df)
        st.session_state.df_raw.update(updated_df)
        st.rerun()

    # (5~6단계 및 기타 요약 로직은 동일)
