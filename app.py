import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import re
import unicodedata
import gspread
from datetime import datetime, timedelta, timezone
import io

# 1. 환경 설정
KST = timezone(timedelta(hours=9))
st.set_page_config(layout="wide", page_title="저스트원 v9.2")

# [새로고침 방지]
components.html("<script>window.onbeforeunload = function() { return '변경사항이 저장되지 않을 수 있습니다.'; };</script>", height=0)

# --- [공통 함수] ---
def super_clean(t):
    if not t: return ""
    t = unicodedata.normalize('NFC', str(t))
    return re.sub(r'[^a-zA-Z0-9가-힣]', '', t).upper().strip()

def to_i(v):
    try: 
        val = str(v).replace(",", "").strip()
        return int(float(val)) if val else 0
    except: return 0

def get_sheet():
    try:
        client = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        return client.open_by_key("1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg")
    except Exception as e:
        st.error(f"📡 시트 연결 실패: {e}")
        return None

def auto_idx(cols, keys, exclude_keys=None):
    for i, c in enumerate(cols):
        c_str = str(c).upper().replace(" ", "")
        if exclude_keys and any(k.upper() in c_str for k in exclude_keys): continue
        if any(k.upper() in c_str for k in keys): return i
    return 0

# --- [공통 데이터 로드] ---
# 여기서부터는 왼쪽 벽에 딱 붙여야 합니다! (공백 제거)
st.divider()
sh = get_sheet()

if sh:
    ws_log = sh.worksheet("발주기록")
    raw_logs = ws_log.get_all_values()
    
    # 데이터가 있는 경우에만 아래 단계들 실행
    if len(raw_logs) > 1:
        df_logs = pd.DataFrame(raw_logs[1:], columns=[c.strip() for c in raw_logs[0]])
        d_col = next((c for c in df_logs.columns if '날짜' in c), df_logs.columns[0])
        v_col = next((c for c in df_logs.columns if '공급처' in c), None)

        # 날짜/시간 전처리 (공통)
        df_logs['pure_dt'] = df_logs[d_col].str.strip()
        df_logs['pure_date'] = df_logs['pure_dt'].str.split(' ').str[0]
        df_logs['pure_time'] = df_logs['pure_dt'].str.split(' ').str[1].str[:5]


# --- [메인 화면] ---
st.title("📦 저스트원 통합 재고 관리")

# 1단계: 업로드
st.header("1️⃣ 데이터 업로드")
if 'reset_trigger' not in st.session_state: st.session_state.reset_trigger = 0

col_up1, col_up2 = st.columns([8, 2])
with col_up1:
    up_file = st.file_uploader("📂 파일 업로드", type=['xlsx', 'xls', 'csv'], key=f"uploader_{st.session_state.reset_trigger}")
with col_up2:
    st.write("") 
    st.write("") 
    if st.button("🔄 전체 초기화", use_container_width=True):
        for k in list(st.session_state.keys()): del st.session_state[k]
        st.rerun()

if up_file:
    if 'df_raw' not in st.session_state:
        try:
            df = pd.read_csv(up_file) if up_file.name.endswith('.csv') else pd.read_excel(up_file)
            df.columns = [str(c).strip() for c in df.columns]
            st.session_state.df_raw = df.fillna("")
            sh = get_sheet()
            if sh:
                ws = sh.worksheet("발주기록")
                all_vals = ws.get_all_values()
                r_map = {}
                if len(all_vals) > 1:
                    for row in all_vals[1:]:
                        if len(row) < 3: continue
                        key = super_clean(row[1]) + super_clean(row[2])
                        r_map[key] = r_map.get(key, 0) + (to_i(row[5]) + to_i(row[6]))
                st.session_state.r_map = r_map
        except Exception as e:
            st.error(f"⚠️ 파일 로드 실패: {e}")
            st.stop()

    cols = st.session_state.df_raw.columns.tolist()
    st.divider()
    
    # 2단계(매핑) & 3단계(설정) - 5:5 비율 유지
    col_step2, col_step3 = st.columns([1, 1])
    with col_step2:
        st.header("2️⃣ 필드 매핑")
        cl, cr = st.columns(2)
        with cl:
            s_so = st.selectbox("🚫 품절 여부", cols, index=auto_idx(cols, ['품절']), key="so_box")
            s_vn = st.selectbox("🏭 공급처", cols, index=auto_idx(cols, ['공급처']), key="vn_box")
            s_vi = st.selectbox("🆔 공급처 상품명", cols, index=auto_idx(cols, ['공급처상품명']), key="vi_box")
            s_it = st.selectbox("📦 상품명", cols, index=auto_idx(cols, ['상품명']), key="it_box")
            s_op = st.selectbox("🎨 옵션", cols, index=auto_idx(cols, ['옵션']), key="op_box")
        with cr:
            s_rd = st.selectbox("📅 등록일", cols, index=auto_idx(cols, ['등록일']), key="rd_box")
            s_st = st.selectbox("🏢 정상재고", cols, index=auto_idx(cols, ['정상재고']), key="st_box")
            s_av = st.selectbox("✅ 가용재고", cols, index=auto_idx(cols, ['가용재고']), key="av_box")
            s_t3 = st.selectbox("🔥 3일 발주합계", cols, index=auto_idx(cols, ['3일']), key="t3_box")
            s_t7 = st.selectbox("📅 7일 발주합계", cols, index=auto_idx(cols, ['7일', '1주']), key="t7_box")

    with col_step3:
        st.header("3️⃣ 수치 설정")
        lt = st.number_input("⏳ 리드타임 (입고 대기)", value=7, key="lt_val")
        ss = st.number_input("🛡️ 안전재고 (최소 유지)", value=3, key="ss_val")
        st.write("")
        if st.button("📊 분석 실행", type="primary", use_container_width=True):
            df = st.session_state.df_raw.copy()
            for c in [s_av, s_t3, s_t7]: df[c] = df[c].apply(to_i)
            df['clean_key'] = df.apply(lambda r: super_clean(r[s_it]) + super_clean(r[s_op]), axis=1)
            df['리오더 수량'] = df['clean_key'].map(st.session_state.r_map).fillna(0).astype(int)
            df['1일 판매량'] = df.apply(lambda r: int(round(r[s_t7]/7)) if r[s_t7]>0 else (int(round(r[s_t3]/3)) if r[s_t3]>0 else 0), axis=1)
            df['권장발주수량'] = ((df['1일 판매량'] * (lt + ss)) - (df[s_av] + df['리오더 수량'])).clip(lower=0).astype(int)
            df['상태'] = df['권장발주수량'].apply(lambda x: "🚨 긴급" if x > 0 else "✅ 정상")
            urgent_items = df[df['권장발주수량'] > 0][s_it].unique()
            df['item_urgent_group'] = df[s_it].isin(urgent_items)
            df['입고차감'] = 0
            df['추가발주'] = 0
            df['메모'] = ""
            st.session_state.analyzed_data = df.sort_values(by=['item_urgent_group', s_it, s_op], ascending=[False, True, True])
            st.session_state.final_mapping = {'vn':s_vn, 'it':s_it, 'op':s_op, 'vi':s_vi, 'av':s_av, 't3':s_t3, 'so':s_so}

    # 4~5단계: 발주 편집 및 저장
    if 'analyzed_data' in st.session_state:
        st.divider()
        st.header("4️⃣~5️⃣ 발주 편집 및 저장")
        m = st.session_state.final_mapping
        
        f1, f2 = st.columns([3, 2])
        with f1:
            f_type = st.radio("🔍 필터", ["전체", "🚨 긴급(묶음)", "✅ 정상", "🚫 품절"], horizontal=True)
        with f2:
            search_q = st.text_input("🔎 상품명 검색")

        df_view = st.session_state.analyzed_data.copy()
        if f_type == "🚨 긴급(묶음)": df_view = df_view[df_view['item_urgent_group'] == True]
        elif f_type == "✅ 정상": df_view = df_view[df_view[m['so']].astype(str).str.contains("정상", na=False)]
        elif f_type == "🚫 품절": df_view = df_view[df_view[m['so']].astype(str).str.contains("품절", na=False)]
        if search_q: df_view = df_view[df_view[m['it']].str.contains(search_q, case=False, na=False)]

        t_cols = ['상태', m['vn'], m['it'], m['op'], m['vi'], m['av'], '리오더 수량', '입고차감', '추가발주', m['t3'], '1일 판매량', '권장발주수량', '메모']
        edited_df = st.data_editor(df_view[t_cols], use_container_width=True, hide_index=True, key="main_editor")

        if st.button("💾 일괄 저장", type="primary", use_container_width=True):
            to_save = edited_df[(edited_df['입고차감'] != 0) | (edited_df['추가발주'] > 0)]
            if not to_save.empty:
                try:
                    sh = get_sheet()
                    ws_main = sh.worksheet("발주기록")
                    now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M')
                    
                    # 저장할 행 생성
                    rows = [[
                        now_s, 
                        str(r[m['it']]), 
                        str(r[m['op']]), 
                        str(r[m['vi']]), 
                        0, 
                        int(r['리오더 수량']), 
                        int(r['추가발주']) - int(r['입고차감']), # 이 값이 리오더 수량에 합산됨
                        int(r['권장발주수량']), 
                        str(r['메모']), 
                        str(r[m['vn']])
                    ] for _, r in to_save.iterrows()]
                    
                    ws_main.append_rows(rows)
                    st.success("✅ 구글 시트에 성공적으로 저장되었습니다!")
                    
                    # 🔥 [핵심] 저장 후 리오더 수량(r_map)을 즉시 다시 계산하기 위해 세션 비우기
                    if 'r_map' in st.session_state:
                        del st.session_state.r_map
                    
                    # 화면을 새로고침하여 상단 분석 수치에 즉시 반영
                    st.rerun() 
                    
                except Exception as e:
                    st.error(f"❌ 저장 중 오류 발생: {e}")
            else:
                st.warning("⚠️ 저장할 변경 내역(입고차감 또는 추가발주)이 없습니다.")

        

# ------------------------------------------------------------------
# 6️⃣단계: 저장 내역 상세 검색 (독립 블록)
# ------------------------------------------------------------------
st.divider()
st.header("6️⃣ 저장 내역 상세 검색")

# 상단 필터 레이아웃
c1, c2, c3 = st.columns([1, 1, 1])
with c1:
    q_date_6 = st.date_input("📅 조회 날짜 선택", value=datetime.now(KST).date(), key="date_6_v9")
with c2:
    st.write("") # 간격 맞춤
    btn_load_6 = st.button("🚀 데이터 조회하기", use_container_width=True, type="primary", key="btn_6_v9")

if btn_load_6:
    sh = get_sheet()
    if sh:
        ws = sh.worksheet("발주기록")
        raw = ws.get_all_values()
        if len(raw) > 1:
            df_6 = pd.DataFrame(raw[1:], columns=[c.strip() for c in raw[0]])
            
            # 날짜 필터링을 위한 전처리
            df_6['pure_date'] = df_6.iloc[:, 0].str.split(' ').str[0]
            target_s = q_date_6.strftime('%Y-%m-%d')
            res_6 = df_6[df_6['pure_date'] == target_s].copy()
            
            if not res_6.empty:
                # [사장님 요청 컬럼 순서 정리]
                # 날짜 => 업체명 => 상품명 => 옵션 => 공급처상품명 => 가용재고 => 기존리오더 => 입고수량 => 추가발주 => 권장발주 => 메모
                # 시트의 실제 헤더 명칭과 똑같아야 합니다. (G열은 상황에 따라 '수량' 혹은 '입고수량'으로 표시)
                
                # 1. 시트의 컬럼명을 유연하게 매칭
                d_col = res_6.columns[0]  # 날짜
                v_col = res_6.columns[-1] # 업체명(공급처)
                qty_col = res_6.columns[6] # G열 (입고수량/추가발주 합산값)
                
                ordered_cols = [
                    d_col, v_col, "상품명", "옵션", "공급처상품명", 
                    "가용재고", "기존리오더", qty_col, "추가발주", "권장발주수량", "메모"
                ]
                
                # 실제 존재하는 컬럼만 필터링 (에러 방지 및 끝부분 보조컬럼 미노출)
                final_cols = [c for c in ordered_cols if c in res_6.columns]
                
                st.write(f"✅ {target_s} 내역 조회 결과 (최신순)")
                st.dataframe(
                    res_6[final_cols].iloc[::-1], 
                    use_container_width=True, 
                    hide_index=True
                )
                
                csv_6 = res_6[final_cols].to_csv(index=False).encode('utf-8-sig')
                st.download_button("📥 현재 결과 CSV 다운로드", data=csv_6, file_name=f"발주내역_{target_s}.csv", use_container_width=True)
            else:
                st.info("해당 날짜에 저장된 데이터가 없습니다.")

# ------------------------------------------------------------------
# 7️⃣단계: 실시간 리오더 최종 잔량 상황판 (독립 블록)
# ------------------------------------------------------------------
st.divider()
st.header("7️⃣ 실시간 리오더 최종 잔량 상황판")

c1_7, c2_7 = st.columns([2, 1])
with c1_7:
    st.info("💡 '현황판 업데이트' 버튼을 누르면 전체 기간의 미입고 잔량을 집계합니다. (합산 잔량이 0 이하인 품목은 자동 제외)")
with c2_7:
    btn_load_7 = st.button("📊 현황판 업데이트", use_container_width=True, key="btn_7_v9", type="secondary")

if btn_load_7:
    sh = get_sheet()
    if sh:
        ws = sh.worksheet("발주기록")
        raw = ws.get_all_values()
        if len(raw) > 1:
            df_7 = pd.DataFrame(raw[1:], columns=[c.strip() for c in raw[0]])
            
            # 수량(G열)과 기존리오더(F열) 수치화
            df_7['qty_val'] = df_7.iloc[:, 6].apply(to_i) # 추가발주 및 차감액
            v_col_7 = df_7.columns[-1] # 업체명
            
            # 상세 그룹화 (업체/상품/옵션별 잔량 합산)
            # 여기서는 상품명(index 1), 옵션(index 2) 등을 사용
            df_sum = df_7.groupby([v_col_7, df_7.columns[1], df_7.columns[2]], as_index=False)['qty_val'].sum()
            
            # 잔량이 0보다 큰(미입고) 데이터만 추출
            df_remain = df_sum[df_sum['qty_val'] > 0].copy()
            df_remain.columns = ['업체명', '상품명', '옵션', '미입고잔량']
            
            if not df_remain.empty:
                # 업체별 총량 대시보드
                st.write("### 🏭 업체별 미입고 합계")
                v_metrics = df_remain.groupby('업체명')['미입고잔량'].sum().reset_index()
                m_cols = st.columns(4)
                for idx, r in enumerate(v_metrics.itertuples()):
                    with m_cols[idx % 4]:
                        st.metric(r.업체명, f"{int(r.미입고잔량):,} 개")
                
                st.write("#### 📋 상세 미입고 리스트")
                st.dataframe(df_remain.sort_values('미입고잔량', ascending=False), use_container_width=True, hide_index=True)
            else:
                st.success("🎉 현재 모든 리오더가 입고 완료되어 잔량이 없습니다!")
