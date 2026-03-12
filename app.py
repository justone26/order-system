import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import holidays

st.set_page_config(layout="wide", page_title="재고 관리 시스템")

# [디자인] 제목과 버튼 분리
st.markdown("""
    <div style="font-size: 55px; font-weight: 900; color: #000; margin-bottom: 20px;">
        📦 재고 관리 및 발주 시스템
    </div>
""", unsafe_allow_html=True)

if st.button("🔄 시스템 초기화"):
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

# [핵심] 가용재고 증가 시 입고예정 자동 차감 함수
def auto_reduce_reorder(df, edited_df):
    # 가용재고 변경분 계산
    diff = edited_df['가용재고'] - df['가용재고']
    
    # 가용재고가 늘어난 경우(입고 완료), 1차 -> 2차 순서로 차감
    for i in diff[diff > 0].index:
        amount_to_reduce = diff[i]
        
        # 1차에서 먼저 차감
        reduce_1st = min(edited_df.at[i, '1차 입고예정'], amount_to_reduce)
        edited_df.at[i, '1차 입고예정'] -= reduce_1st
        amount_to_reduce -= reduce_1st
        
        # 남은 양이 있으면 2차에서 차감
        if amount_to_reduce > 0:
            reduce_2nd = min(edited_df.at[i, '2차 입고예정'], amount_to_reduce)
            edited_df.at[i, '2차 입고예정'] -= reduce_2nd
            
    return edited_df

# ... (이후 4단계 데이터 편집 로직 수정)

    # 4단계: 검색 및 데이터 편집
    st.subheader("📊 4단계: 검색 및 데이터 편집")
    df_final = st.session_state.df_raw[edit_cols]
    
    # 편집기 실행
    edited_df = st.data_editor(df_final, use_container_width=True, disabled=disabled_cols)
    
    # [핵심] 편집이 완료되면 자동 차감 로직 실행
    updated_df = auto_reduce_reorder(st.session_state.df_raw, edited_df)
    st.session_state.df_raw.update(updated_df)
