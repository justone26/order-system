import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="재고 관리 및 발주 시스템", layout="wide")
st.title("📦 재고 관리 및 발주 시스템")

# [세션 상태 초기화] 데이터를 고정 저장하기 위함
if 'df_processed' not in st.session_state:
    st.session_state.df_processed = None

uploaded_file = st.file_uploader("엑셀 또는 CSV 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

if uploaded_file is not None:
    # 1. 파일 로드 및 초기 설정
    if st.session_state.df_processed is None:
        try:
            if uploaded_file.name.endswith('.csv'):
                try: df = pd.read_csv(uploaded_file, encoding='utf-8')
                except: df = pd.read_csv(uploaded_file, encoding='cp949')
            elif uploaded_file.name.endswith('.xls'):
                df = pd.read_excel(uploaded_file, engine='xlrd')
            else:
                df = pd.read_excel(uploaded_file, engine='openpyxl')
            
            df = df.loc[:, ~df.columns.duplicated()]
            # 입고예정수량 기본값 생성
            if "입고예정수량(리오더)" not in df.columns:
                df["입고예정수량(리오더)"] = 0
            
            st.session_state.df_processed = df
        except Exception as e:
            st.error(f"파일 로드 오류: {e}")

if st.session_state.df_processed is not None:
    df = st.session_state.df_processed
    columns = df.columns.tolist()

    # 매핑 섹션
    st.subheader("⚙️ 1단계: 컬럼 매핑")
    # (매핑 로직은 동일)
    
    # 4. 분석 실행
    if st.button("🚀 분석 실행"):
        # 데이터 정제 및 계산
        df['일일 판매량(기준)'] = (pd.to_numeric(df.get('3일 발주 합계', 0), errors='coerce') / 3).fillna(0).astype(int)
        df['권장 발주량'] = (df['일일 판매량(기준)'] * 3 - (pd.to_numeric(df.get('가용재고', 0), errors='coerce') + df["입고예정수량(리오더)"])).clip(lower=0).astype(int)
        
        st.session_state.df_processed = df

    # [수정] 데이터 에디터를 항상 보여주어 데이터가 사라지지 않게 함
    st.subheader("📊 편집 및 결과 확인")
    
    # 사용자가 데이터를 수정하면 세션 상태에 즉시 업데이트
    edited_df = st.data_editor(st.session_state.df_processed, use_container_width=True)
    st.session_state.df_processed = edited_df

    # 다운로드 버튼
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        edited_df.to_excel(writer, index=False)
    
    st.download_button("📥 최종 결과 엑셀 다운로드", data=buffer.getvalue(), file_name="최종_발주서.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
