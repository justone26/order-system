import streamlit as st
import pandas as pd
import numpy as np
import re
import unicodedata
import gspread
from datetime import datetime, timedelta, timezone

# 1. 환경 설정
KST = timezone(timedelta(hours=9))
st.set_page_config(layout="wide", page_title="저스트원 발주 시스템")

# --- [필수 함수] ---
def find_idx(cols, keys):
    for i, c in enumerate(cols):
        if any(k in str(c) for k in keys): return i
    return 0

def super_clean(t):
    if not t: return ""
    t = unicodedata.normalize('NFC', str(t))
    return re.sub(r'[^a-zA-Z0-9가-힣]', '', t).upper().strip()

def to_i(v):
    try: return int(float(str(v).replace(",", "")))
    except: return 0

def get_sheet():
    try:
        client = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        return client.open_by_key("1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg")
    except: return None

# --- [메인 로직 시작] ---
st.header("1️⃣ 파일 업로드")
up_file = st.file_uploader("엑셀 업로드", type=['xlsx', 'xls'])

if up_file:
    if 'df_raw' not in st.session_state:
        st.session_state.df_raw = pd.read_excel(up_file)
        st.session_state.analyzed = False

    df_work = st.session_state.df_raw
    cols = df_work.columns.tolist()

    # --- [2단계: 매핑 설정 - 사장님 틀 유지] ---
    st.subheader("⚙️ 2단계: 매핑 설정")
    c1, c2 = st.columns(2)
    sold_out = c1.selectbox("품절 여부", cols, index=find_idx(cols, ['품절']))
    vendor = c1.selectbox("공급처", cols, index=find_idx(cols, ['공급처']))
    item = c1.selectbox("상품명", cols, index=find_idx(cols, ['상품명']))
    option = c1.selectbox("옵션", cols, index=find_idx(cols, ['옵션']))
    vendor_item = c1.selectbox("공급처 상품명", cols, index=find_idx(cols, ['공급처상품명']))

    reg_date = c2.selectbox("등록일", cols, index=find_idx(cols, ['등록일']))
    stock = c2.selectbox("정상재고", cols, index=find_idx(cols, ['정상재고']))
    avail = c2.selectbox("가용재고", cols, index=find_idx(cols, ['가용재고']))
    t3day = c2.selectbox("3일 발주합계", cols, index=find_idx(cols, ['3일']))
    t1week = c2.selectbox("7일 발주합계", cols, index=find_idx(cols, ['7일', '1주']))

# --- [3단계: 분석 설정 및 실행] ---
    st.subheader("⚙️ 3단계: 분석 설정")
    col_lt, col_ss = st.columns(2)
    lead_time = col_lt.number_input("리드타임 (일)", value=10, key="v3_lt_final")
    safety_stock = col_ss.number_input("안전재고 (일 수)", value=7, key="v3_ss_final")

    if st.button("🚀 분석 실행", type="primary", use_container_width=True):
        # 1. 원본 데이터 복사
        df = st.session_state.df_raw.copy()
        
        # 2. 구글 시트에서 기존 리오더 수량 가져오기 (에러 방지 로직)
        sh = get_sheet()
        r_map = {}
        if sh:
            try:
                ws = sh.worksheet("발주기록")
                logs = ws.get_all_values()
                
                if len(logs) > 1: # 데이터가 있을 때만
                    df_l = pd.DataFrame(logs[1:], columns=[c.strip() for c in logs[0]])
                    
                    # iloc를 사용하여 열 순서로 안전하게 접근 (1:상품명, 2:옵션, 6:변동수량)
                    if df_l.shape[1] >= 7:
                        df_l['k'] = df_l.apply(lambda r: super_clean(r.iloc[1]) + super_clean(r.iloc[2]), axis=1)
                        # G열(인덱스 6) 수량을 합산하여 리오더 맵 생성
                        r_map = df_l.groupby('k').apply(lambda x: x.iloc[:, 6].apply(to_i).sum()).to_dict()
                else:
                    st.info("💡 발주기록이 비어있어 기존 리오더를 0으로 시작합니다.")
            except Exception as e:
                st.warning(f"⚠️ 시트 연동 중 알림: {e}")

        # 3. 데이터 가공 및 수식 적용
        # 상품명+옵션으로 고유 키 생성
        df['clean_k'] = df.apply(lambda r: super_clean(r[item]) + super_clean(r[option]), axis=1)
        
        # 기존 리오더 매칭
        df['기존리오더'] = df['clean_k'].map(r_map).fillna(0).astype(int)
        
        # 일평균 판매량 계산 (7일 우선, 없으면 3일)
        avail_val = pd.to_numeric(df[avail], errors='coerce').fillna(0)
        t1w_val = pd.to_numeric(df[t1week], errors='coerce').fillna(0)
        t3d_val = pd.to_numeric(df[t3day], errors='coerce').fillna(0)
        
        # 판매량 기반 권장 발주량 계산
        daily_avg = t1w_val / 7
        df['권장발주수량'] = ((daily_avg * lead_time) + (daily_avg * safety_stock) - (avail_val + df['기존리오더'])).clip(lower=0).astype(int)
        
        # 편집을 위한 기본 열 생성
        df['추가발주'] = 0
        df['입고차감'] = 0
        df['메모'] = ""
        
        # 정렬 로직: 권장발주가 있는(🚨긴급) 품목을 위로
        df['상태'] = df['권장발주수량'].apply(lambda x: "🚨 긴급" if x > 0 else "✅ 정상")
        df = df.sort_values(by=['상태', item], ascending=[False, True])

        # 4. 결과 저장 및 화면 갱신
        st.session_state.df_raw = df
        st.session_state.analyzed = True
        st.success("📊 분석이 완료되었습니다! 아래에서 확인하세요.")
        st.rerun()



    
   # 3단계 분석 완료 후 나타나는 구역
    if st.session_state.get('analyzed'):
        st.divider()
        st.header("4️⃣~5️⃣ 발주 수량 편집 및 저장")
        
        # 분석된 전체 데이터 가져오기
        df_final = st.session_state.df_raw.copy()

        # --- [필터 검색 바 영역] ---
        st.subheader("🔍 리스트 필터링")
        f_c1, f_c2, f_c3 = st.columns([1, 1, 2])
        
        with f_c1:
            # 업체명 필터 (2단계에서 고른 vendor 컬럼 사용)
            v_list = ["전체"] + sorted(df_final[vendor].unique().tolist())
            sel_v = st.selectbox(f"🏭 {vendor} 선택", v_list, key="filter_v")
            
        with f_c2:
            # 상태 필터 (긴급/정상)
            s_list = ["전체", "🚨 긴급", "✅ 정상"]
            sel_s = st.selectbox("🚦 상태 필터", s_list, key="filter_s")
            
        with f_c3:
            # 상품명 검색 (2단계에서 고른 item 컬럼 사용)
            q_word = st.text_input(f"🔎 {item} 검색", placeholder="검색어를 입력하세요...", key="filter_q")

        # --- [데이터 필터링 적용] ---
        disp_df = df_final.copy()
        if sel_v != "전체":
            disp_df = disp_df[disp_df[vendor] == sel_v]
        if sel_s != "전체":
            disp_df = disp_df[disp_df['상태'] == sel_s]
        if q_word:
            disp_df = disp_df[disp_df[item].str.contains(q_word, case=False, na=False)]

        # --- [편집기 실행] ---
        # 사장님이 요청한 항목 순서로 열 배치
        display_cols = [
            '상태', vendor, item, option, vendor_item, avail, 
            '기존리오더', '권장발주수량', '추가발주', '입고차감', '메모'
        ]
        safe_cols = [c for c in display_cols if c in disp_df.columns]

        st.info(f"💡 현재 {len(disp_df)}개의 항목이 표시되고 있습니다.")
        
        edited_df = st.data_editor(
            disp_df[safe_cols],
            column_config={
                "상태": st.column_config.TextColumn("상태", width="small"),
                vendor: st.column_config.TextColumn(vendor, width="medium"),
                item: st.column_config.TextColumn(item, width="large"),
                "추가발주": st.column_config.NumberColumn("➕ 추가발주", step=1),
                "입고차감": st.column_config.NumberColumn("➖ 입고차감", step=1),
                "메모": st.column_config.TextColumn("📝 메모", width="large")
            },
            disabled=[c for c in safe_cols if c not in ['추가발주', '입고차감', '메모']],
            hide_index=True,
            use_container_width=True,
            key="v12_editor"
        )

        # --- [5단계: 일괄 저장] ---
        if st.button("💾 편집된 내역 구글 시트로 일괄 저장", type="primary", use_container_width=True):
            # 수정한 내역이 있는 행만 추출
            to_save = edited_df[(edited_df['추가발주'] > 0) | (edited_df['입고차감'] > 0)]
            
            if not to_save.empty:
                try:
                    sh = get_sheet()
                    ws = sh.worksheet("발주기록")
                    now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M')
                    
                    rows = []
                    for _, r in to_save.iterrows():
                        rows.append([
                            now_s,
                            str(r[item]),
                            str(r[option]),
                            str(r[vendor_item]),
                            int(to_i(r[avail])),
                            int(r['기존리오더']),
                            int(r['추가발주']) - int(r['입고차감']), # 변동 수량
                            int(r['권장발주수량']),
                            str(r['메모']),
                            str(r[vendor])
                        ])
                    
                    ws.append_rows(rows)
                    st.success(f"✅ {len(rows)}건 저장 성공! 수치 갱신을 위해 재실행합니다.")
                    
                    # 수치 갱신을 위해 분석 상태 초기화 후 재실행
                    st.session_state.analyzed = False
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ 저장 오류: {e}")
            else:
                st.warning("⚠️ 저장할 변경 사항이 없습니다.")

        
        # 6️⃣단계: 저장 내역 상세 검색
        st.divider()
        st.subheader("6️⃣ 저장 내역 상세 검색")
        c6_1, c6_2 = st.columns(2)
        q_date = c6_1.date_input("날짜 선택")
        if c6_2.button("🚀 검색 실행", use_container_width=True):
            sh = get_sheet()
            if sh:
                raw_logs = sh.worksheet("발주기록").get_all_values()
                df_log = pd.DataFrame(raw_logs[1:], columns=raw_logs[0])
                target = q_date.strftime('%Y-%m-%d')
                res = df_log[df_log.iloc[:, 0].str.contains(target)].copy()
                
                # 순서 고정: 날짜, 업체명, 상품명, 옵션, 공급처상품명, 가용, 기존, 수량(G), 추가, 권장, 메모
                # (중복 제거 로직 포함)
                col_view = [res.columns[0], res.columns[9], res.columns[1], res.columns[2], res.columns[3], 
                            res.columns[4], res.columns[5], res.columns[6], res.columns[7], res.columns[8]]
                st.dataframe(res[col_view].iloc[::-1], use_container_width=True, hide_index=True)

        # 7️⃣단계: 실시간 잔량 상황판
        st.divider()
        st.subheader("7️⃣ 실시간 리오더 최종 잔량 상황판")
        if st.button("📊 현황판 업데이트"):
            sh = get_sheet()
            if sh:
                raw = sh.worksheet("발주기록").get_all_values()
                df_7 = pd.DataFrame(raw[1:], columns=raw[0])
                df_7['qty'] = df_7.iloc[:, 6].apply(to_i)
                v_sum = df_7.groupby(df_7.columns[9])['qty'].sum().reset_index()
                v_sum = v_sum[v_sum['qty'] > 0]
                
                m_cols = st.columns(4)
                for i, r in enumerate(v_sum.itertuples()):
                    with m_cols[i % 4]: st.metric(r[1], f"{int(r[2])} 개")
