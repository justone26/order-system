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

        

# ---------------------------------------------------------
                # 6️⃣단계: 저장 내역 상세 검색 (컬럼 순서 정리 버전)
                # ---------------------------------------------------------
                st.header("6️⃣ 저장 내역 상세 검색")
                
                s_col1, s_col2, s_col3 = st.columns([1, 1.5, 1.5])
                with s_col1:
                    q_date = st.date_input("📅 날짜 선택", value=datetime.now(KST).date(), key="s_date_6")
                    target_date = q_date.strftime('%Y-%m-%d') if q_date else ""
                
                f_logs_6 = df_logs[df_logs['pure_date'] == target_date].copy()

                with s_col3:
                    times_6 = sorted(f_logs_6['pure_time'].dropna().unique(), reverse=True)
                    q_time_6 = st.selectbox(f"⏰ 저장 회차 ({len(times_6)}회)", ["전체 보기"] + times_6, key="s_time_6")

                with s_col2:
                    q_item_6 = st.text_input("🔎 상품명 검색", key="s_item_6")

                # 필터링 적용
                df_res_6 = f_logs_6.copy()
                if q_time_6 != "전체 보기":
                    df_res_6 = df_res_6[df_res_6['pure_time'] == q_time_6]
                if q_item_6:
                    i_col = next((c for c in df_res_6.columns if '상품명' in c), None)
                    if i_col:
                        df_res_6 = df_res_6[df_res_6[i_col].str.contains(q_item_6, case=False)]

                # 🔥 [핵심] 사장님이 요청하신 순서대로 컬럼 재배치 및 미노출 설정
                # 기존 시트의 컬럼 순서를 무시하고 아래 정의한 순서대로만 화면에 출력합니다.
                display_cols = [
                    d_col,          # 날짜(발주시간)
                    v_col,          # 업체명(공급처)
                    "상품명", 
                    "옵션", 
                    "공급처상품명", 
                    "가용재고", 
                    "기존리오더", 
                    "수량",         # 입고수량/발주수량(G열)
                    "추가발주",     # 실제 저장 시 계산된 추가발주
                    "권장발주수량", # 당시 계산된 권장량
                    "메모"
                ]
                
                # 존재하는 컬럼만 필터링 (에러 방지)
                actual_display_cols = [c for c in display_cols if c in df_res_6.columns]
                
                # 표 출력 (보조 컬럼 pure_dt, pure_date, pure_time 등은 자동 미노출됨)
                st.dataframe(
                    df_res_6[actual_display_cols].iloc[::-1], 
                    use_container_width=True, 
                    hide_index=True
                )
                
                if not df_res_6.empty:
                    csv_6 = df_res_6[actual_display_cols].to_csv(index=False).encode('utf-8-sig')
                    st.download_button("📥 현재 검색 결과 다운로드", data=csv_6, file_name=f"발주내역_{target_date}.csv", use_container_width=True)


# ---------------------------------------------------------
                # 7️⃣단계: 실시간 리오더 최종 잔량 상황판 (제로화/메모 보정 통합)
                # ---------------------------------------------------------
                st.divider()
                st.subheader("🚀 7️⃣단계: 실시간 리오더 최종 잔량 상황판")

                # 1. 데이터 클리닝 및 전처리
                df_v7 = df_logs.copy()
                
                # 수치 변환 (G열: 추가발주/입고차감 수량, F열: 기존리오더 수량)
                # 시트 구조에 따라 iloc 인덱스는 확인이 필요할 수 있습니다.
                df_v7["기존리오더"] = df_v7.iloc[:, 5].apply(to_i)
                df_v7["추가발주"] = df_v7.iloc[:, 6].apply(to_i)
                df_v7["최종잔량"] = df_v7["기존리오더"] + df_v7["추가발주"]
                
                # 메모 보정: "입고차감"만 적힌 경우 수량을 붙여줌
                def fix_memo_v7(row):
                    m = str(row.get('메모', '')).strip()
                    q = row['추가발주']
                    if q < 0 and (m == "입고차감" or m == ""):
                        return f"{abs(q)}개 입고차감"
                    return m
                
                if '메모' in df_v7.columns:
                    df_v7["메모_보정"] = df_v7.apply(fix_memo_v7, axis=1)
                else:
                    df_v7["메모_보정"] = df_v7["추가발주"].apply(lambda x: f"{abs(x)}개 입고차감" if x < 0 else "")

                # 2. 필터 UI (기존 기능 유지)
                f_col1, f_col2, f_col3 = st.columns([1, 1, 2])
                with f_col1:
                    d_range = st.date_input("🗓️ 조회 기간", 
                                            value=((datetime.now(KST)-timedelta(days=30)).date(), datetime.now(KST).date()), 
                                            key="v7_date_range")
                with f_col2:
                    v_choice = st.selectbox("🏭 업체 선택", ["전체 업체"] + sorted(df_v7[v_col].unique().tolist()), key="v7_vendor_sel")
                with f_col3:
                    q_v7 = st.text_input("🔍 상품명/옵션 검색", key="v7_search_input")

                # 3. 필터링 로직
                df_f = df_v7.copy()
                if isinstance(d_range, (list, tuple)) and len(d_range) == 2:
                    df_f = df_f[(df_f["pure_date"] >= d_range[0].strftime('%Y-%m-%d')) & 
                                (df_f["pure_date"] <= d_range[1].strftime('%Y-%m-%d'))]
                
                if v_choice != "전체 업체":
                    df_f = df_f[df_f[v_col] == v_choice]
                
                if q_v7:
                    i_col = next((c for c in df_f.columns if '상품명' in c), None)
                    o_col = next((c for c in df_f.columns if '옵션' in c), None)
                    cond = df_f[i_col].str.contains(q_v7, case=False) if i_col else False
                    if o_col: cond |= df_f[o_col].str.contains(q_v7, case=False)
                    df_f = df_f[cond]

                # 4. 상세 리스트 그룹화 및 제로화(Clip)
                if not df_f.empty:
                    it_col = next((c for c in df_f.columns if '상품명' in c), "상품명")
                    op_col = next((c for c in df_f.columns if '옵션' in c), "옵션")
                    vi_col = next((c for c in df_f.columns if '공급처상품명' in c), "공급처상품명")
                    
                    df_final = df_f.groupby([v_col, it_col, op_col, vi_col], as_index=False).agg({
                        d_col: "max",
                        "최종잔량": "sum",
                        "메모_보정": lambda x: " / ".join(dict.fromkeys(filter(None, x.astype(str))))
                    })

                    # ⭐ [핵심] 마이너스 잔량은 0으로 처리 (입고가 더 많이 잡힌 경우 방지)
                    df_final["최종잔량"] = df_final["최종잔량"].clip(lower=0)
                    df_final = df_final.sort_values([d_col, v_col], ascending=[False, True])

                    # 5. 업체별 전광판 합산 출력
                    st.write("### 📊 업체별 미입고 현황 (미입고 1개 이상)")
                    df_v_sum = df_final.groupby(v_col)["최종잔량"].sum().reset_index()
                    df_v_sum = df_v_sum[df_v_sum["최종잔량"] > 0].sort_values("최종잔량", ascending=False)
                    
                    if not df_v_sum.empty:
                        v_metrics = st.columns(4)
                        for i, r in enumerate(df_v_sum.itertuples()):
                            with v_metrics[i % 4]:
                                st.metric(label=getattr(r, v_col), value=f"{int(r.최종잔량):,} 개")
                    
                    # 6. 상세 내역 테이블
                    st.write("#### 📋 상세 미입고 리스트")
                    st.dataframe(
                        df_final[df_final["최종잔량"] > 0], # 잔량이 있는 것만 노출
                        use_container_width=True, 
                        hide_index=True,
                        column_config={
                            d_col: st.column_config.TextColumn("🕒 최종발주"),
                            "최종잔량": st.column_config.NumberColumn("🔢 미입고 잔량", format="%d"),
                            "메모_보정": st.column_config.TextColumn("📝 비고(차감내역)", width="large")
                        }
                    )
                else:
                    st.info("조회된 기간 내에 미입고 데이터가 없습니다.")
