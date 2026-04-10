import streamlit as st
import pandas as pd
import numpy as np
import re
import time
import unicodedata
import gspread
import streamlit.components.v1 as components
from datetime import datetime, timedelta, timezone

# 1. 기본 설정 및 한국 시간
KST = timezone(timedelta(hours=9))
st.set_page_config(layout="wide", page_title="저스트원 통합 관리 v4.5")

# --- [세션 상태 초기화] ---
if 'df_raw' not in st.session_state: st.session_state.df_raw = None
if 'analyzed' not in st.session_state: st.session_state.analyzed = False
if 'p' not in st.session_state: st.session_state.p = {}
if 'add_order_dict' not in st.session_state: st.session_state.add_order_dict = {}
if 'upload_key' not in st.session_state: st.session_state.upload_key = 0

# 새로고침 방지
components.html("<script>window.onbeforeunload = function() { return '변경사항이 저장되지 않을 수 있습니다.'; };</script>", height=0)

# --- [공통 보조 함수] ---
def get_sheet():
    try:
        client = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        return client.open_by_key("1uWZ2xeS9Zj5Dpn2zB-enRHNMGGJ8JTl48HfICvVTOdg")
    except Exception as e:
        st.error(f"📡 시트 연결 실패: {e}")
        return None

def super_clean(t):
    if not t: return ""
    t = unicodedata.normalize('NFC', str(t))
    return re.sub(r'[^a-zA-Z0-9가-힣]', '', t).upper().strip()

def to_i(v):
    try: return int(float(str(v).replace(",", "").strip()))
    except: return 0

# [속도 핵심] 실시간 데이터 로드 (분석 시작 시점에 단 1회 호출 권장)
def get_realtime_map():
    try:
        sh = get_sheet()
        if not sh: return {}, {}
        
        # 1. 발주기록 (리오더 합산)
        ws_v = sh.worksheet("발주기록")
        d_v = ws_v.get_all_values()
        r_map = {}
        if len(d_v) > 1:
            for row in d_v[1:]:
                try:
                    key = super_clean(row[1]) + super_clean(row[2])
                    val = to_i(row[5]) # 기존
                    add = to_i(row[6]) # 추가
                    r_map[key] = r_map.get(key, 0) + (val + add)
                except: continue
        return r_map
    except:
        return {}

# --- [메인 UI] ---
st.title("📦 저스트원 통합 재고 관리 (믹스 버전)")
tab1, tab2 = st.tabs(["🏭 제작 상품 관리", "📜 발주 히스토리/상황판"])

with tab1:
    ## 1단계: 업로드
    up_file = st.file_uploader("파일 업로드", type=['xlsx', 'xls', 'csv'], key=f"up_{st.session_state.upload_key}")

    if st.button("🔄 화면 전체 초기화", use_container_width=True):
        for key in list(st.session_state.keys()):
            if key != "upload_key": del st.session_state[key]
        st.session_state.upload_key += 1
        st.rerun()

    if up_file and st.session_state.df_raw is None:
        df = pd.read_csv(up_file) if up_file.name.endswith('.csv') else pd.read_excel(up_file)
        df.columns = [str(c).strip() for c in df.columns]
        st.session_state.df_raw = df.fillna("")
        st.rerun()

    if st.session_state.df_raw is not None:
        ## 2단계: 매핑 (초기 소스 방식의 편리함)
        cols = st.session_state.df_raw.columns.tolist()
        def auto_idx(keys, exclude=None):
            for i, c in enumerate(cols):
                if exclude and any(e in str(c) for e in exclude): continue
                if any(k in str(c) for k in keys): return i
            return 0

        st.subheader("📋 매핑 및 분석 설정")
        c1, c2, c3 = st.columns(3)
        with c1:
            it = st.selectbox("📦 상품명", cols, index=auto_idx(['상품명']))
            op = st.selectbox("🎨 옵션", cols, index=auto_idx(['옵션']))
            vn = st.selectbox("🏭 공급처", cols, index=auto_idx(['공급처']))
        with c2:
            av = st.selectbox("✅ 가용재고", cols, index=auto_idx(['가용재고']))
            t3 = st.selectbox("🔥 3일 판매", cols, index=auto_idx(['3일'], exclude=['7일']))
            t7 = st.selectbox("📅 7일 판매", cols, index=auto_idx(['7일']))
        with c3:
            lt = st.number_input("⏳ 리드타임", value=7)
            ss = st.number_input("🛡️ 안전재고", value=3)
            vi = st.selectbox("🆔 공급처상품명", cols, index=auto_idx(['공급처상품명']))

        if st.button("📊 분석 및 리오더 로드 (속도 최적화)", type="primary", use_container_width=True):
            with st.spinner("📡 실시간 리오더 데이터를 가져오는 중..."):
                r_map = get_realtime_map()
                df = st.session_state.df_raw.copy()
                # 필수 숫자 변환
                for c in [av, t3, t7]: df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0).astype(int)
                
                # 리오더 잔량 매핑
                df['clean_key'] = df.apply(lambda r: super_clean(r[it]) + super_clean(r[op]), axis=1)
                df['기존잔량'] = df['clean_key'].map(r_map).fillna(0).astype(int)
                
                # 계산식 적용
                df['일판매'] = df.apply(lambda r: int(round(r[t7]/7)) if r[t7]>0 else (int(round(r[t3]/3)) if r[t3]>0 else 0), axis=1)
                df['발주권장'] = ((df['일판매'] * (lt + ss)) - (df[av] + df['기존잔량'])).clip(lower=0).astype(int)
                df['상태'] = df['발주권장'].apply(lambda x: "🚨 긴급" if x > 0 else "✅ 정상")
                
                # 초기화
                if '추가발주' not in df.columns: df['추가발주'] = 0
                if '메모' not in df.columns: df['메모'] = ""
                
                st.session_state.df_raw = df
                st.session_state.p = {'it':it, 'op':op, 'vn':vn, 'av':av, 'vi':vi}
                st.session_state.analyzed = True
                st.rerun()

    # --- [4~5단계 통합 편집 영역] ---
    if st.session_state.analyzed:
        st.divider()
        st.subheader("📝 발주 편집 및 리오더 관리")
        
        # 필터링
        f1, f2 = st.columns([1, 2])
        with f1: mode = st.selectbox("🚦 필터", ["긴급만", "전체보기"])
        with f2: query = st.text_input("🔍 검색")
        
        df_edit = st.session_state.df_raw.copy()
        if mode == "긴급만": df_edit = df_edit[df_edit['상태'] == "🚨 긴급"]
        if query: df_edit = df_edit[df_edit[st.session_state.p['it']].astype(str).str.contains(query, case=False)]

        # 에디터 (속도를 위해 on_change 제거, 마지막에 한 번에 저장)
        disp_cols = ['상태', st.session_state.p['vn'], st.session_state.p['it'], st.session_state.p['op'], st.session_state.p['av'], '기존잔량', '추가발주', '발주권장', '메모']
        
        edited_df = st.data_editor(
            df_edit[disp_cols],
            use_container_width=True,
            hide_index=True,
            column_config={
                "추가발주": st.column_config.NumberColumn("➕ 추가발주", min_value=0),
                "메모": st.column_config.TextColumn("📝 메모", width="medium"),
                "기존잔량": st.column_config.NumberColumn("📦 기존", disabled=True),
                "발주권장": st.column_config.NumberColumn("💡 권장", disabled=True),
                "상태": st.column_config.TextColumn("🚦 상태", disabled=True)
            }
        )

        # 저장 로직
        c_save, c_down = st.columns(2)
        with c_save:
            if st.button("💾 구글 시트 최종 저장 (히스토리 기록)", type="primary", use_container_width=True):
                # 편집된 내용 반영 (edited_df -> df_raw)
                # 이 예제에서는 단순화를 위해 edited_df의 추가발주가 0보다 큰 것만 골라 저장합니다.
                to_save = edited_df[edited_df['추가발주'] > 0]
                
                if not to_save.empty:
                    try:
                        sh = get_sheet()
                        ws = sh.worksheet("발주기록")
                        ws_h = sh.worksheet("history")
                        now_s = datetime.now(KST).strftime('%Y-%m-%d %H:%M')
                        
                        rows = []
                        for _, r in to_save.iterrows():
                            rows.append([
                                now_s, str(r[st.session_state.p['it']]), str(r[st.session_state.p['op']]), 
                                "", # 공급처상품명(필요시 추가)
                                int(r[st.session_state.p['av']]), int(r['기존잔량']), int(r['추가발주']), 
                                int(r['발주권장']), str(r['메모']), str(r[st.session_state.p['vn']])
                            ])
                        
                        ws.append_rows(rows)
                        ws_h.append_rows(rows)
                        st.success(f"✅ {len(rows)}건 저장 완료!")
                        time.sleep(1)
                        st.rerun()
                    except Exception as e:
                        st.error(f"저장 실패: {e}")
                else:
                    st.warning("저장할 추가 발주 수량이 없습니다.")

        with c_down:
            # 다운로드 기능
            if not edited_df[edited_df['추가발주'] > 0].empty:
                csv = edited_df[edited_df['추가발주'] > 0].to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
                st.download_button("📥 발주서 다운로드", csv, "발주서.csv", use_container_width=True)


# ==================================================================
# [6단계: 추가 발주 히스토리 관리]
# ==================================================================
if st.session_state.get('analyzed'):
    st.divider()
    st.subheader("📜 6단계: 추가 발주 히스토리 관리")
    st.info("💡 5단계 화면 구성 그대로 조회하되, 메모 열만 제외하여 깔끔하게 보여줍니다.")

    h_f1, h_f2, h_f3 = st.columns([1.5, 1, 2])
    with h_f1:
        v6_date = st.date_input("🗓️ 조회 기간", value=(datetime.now(KST).date(), datetime.now(KST).date()), key="v6_date_pick")
    with h_f2:
        st.write(""); st.write("")
        v6_search_btn = st.button("🔍 내역 조회 실행", use_container_width=True, type="primary")

    if 'v6_storage' not in st.session_state: st.session_state.v6_storage = None

    if v6_search_btn:
        try:
            with st.spinner("📡 history 시트 로드 중..."):
                ws_h = get_sheet().worksheet("history")
                h_all = ws_h.get_all_values()
                if len(h_all) > 1:
                    df_h = pd.DataFrame(h_all[1:], columns=h_all[0])
                    df_h["date_only"] = df_h["발주시간"].astype(str).str.slice(0, 10)
                    start_s = v6_date[0].strftime('%Y-%m-%d')
                    end_s = v6_date[1].strftime('%Y-%m-%d') if len(v6_date) > 1 else start_s
                    
                    # 추가발주가 있는 내역만 필터링
                    q_col = "추가발주량" if "추가발주량" in df_h.columns else "추가발주"
                    df_h[q_col] = pd.to_numeric(df_h[q_col], errors='coerce').fillna(0).astype(int)
                    st.session_state.v6_storage = df_h[(df_h["date_only"] >= start_s) & (df_h["date_only"] <= end_s) & (df_h[q_col] > 0)].copy()
                else: st.session_state.v6_storage = None
        except Exception as e: st.error(f"조회 에러: {e}")

    if st.session_state.v6_storage is not None:
        df_h_view = st.session_state.v6_storage.copy()
        with h_f3:
            h_q = st.text_input("🔍 결과 내 검색 (상품/옵션)", key="v6_inner_search")
            if h_q: df_h_view = df_h_view[df_h_view["상품명"].str.contains(h_q, case=False) | df_h_view["옵션"].str.contains(h_q, case=False)]

        if not df_h_view.empty:
            # ⭐ 사장님 요청: 5단계 화면 구성 그대로 (메모 제외)
            # 순서: 발주시간, 업체명, 상품명, 옵션, 가용재고, 기존리오더, 추가발주
            h_target_q = "추가발주량" if "추가발주량" in df_h_view.columns else "추가발주"
            h_disp_cols = ["발주시간", "업체명", "상품명", "옵션", "가용재고", "기존리오더", h_target_q]
            
            # 숫자형 변환
            for num_c in ["가용재고", "기존리오더", h_target_q]:
                if num_c in df_h_view.columns:
                    df_h_view[num_c] = pd.to_numeric(df_h_view[num_c], errors='coerce').fillna(0).astype(int)

            st.dataframe(
                df_h_view[h_disp_cols].sort_values("발주시간", ascending=False),
                use_container_width=True, hide_index=True,
                column_config={
                    "가용재고": st.column_config.NumberColumn("📦 가용"),
                    "기존리오더": st.column_config.NumberColumn("🗒️ 기존"),
                    h_target_q: st.column_config.NumberColumn("➕ 추가발주")
                }
            )
        else: st.info("조회된 내역이 없습니다.")
        

# ------------------------------------------------------------------
# [7단계: 실시간 상황판] - 업데이트 버튼 클릭 시에만 갱신 (API 최적화)
# ------------------------------------------------------------------
import io
import pandas as pd
import streamlit as st

if st.session_state.get('analyzed'):
    st.divider()
    st.subheader("🚀 7단계: 실시간 리오더 최종 잔량 상황판")
    st.info("💡 [🔄 상황판 데이터 갱신] 버튼을 누를 때만 실시간 데이터를 집계합니다.")

    # [1] 데이터 로드 함수 (캐시 적용)
    @st.cache_data(ttl=600) # 10분간 캐시 유지 (버튼 클릭 시 clear_cache 함)
    def get_v7_minimal_memo():
        try:
            sh = get_sheet()
            # 발주기록
            ws_o = sh.worksheet("발주기록")
            o_all = ws_o.get_all_values()
            o_h_idx = 1 if len(o_all) > 1 and "상품명" in o_all[1] else 0
            df_o = pd.DataFrame(o_all[o_h_idx+1:], columns=o_all[o_h_idx]) if len(o_all) > 1 else pd.DataFrame()
            
            # 입고기록
            ws_r = sh.worksheet("입고기록")
            r_all = ws_r.get_all_values()
            r_h_idx = 1 if len(r_all) > 1 and "상품명" in r_all[1] else 0
            df_r = pd.DataFrame(r_all[r_h_idx+1:], columns=r_all[r_h_idx]) if len(r_all) > 1 else pd.DataFrame()
            
            return df_o, df_r
        except Exception as e:
            st.error(f"시트 연결 실패: {e}")
            return pd.DataFrame(), pd.DataFrame()

    # --- [2. 업데이트 제어 UI] ---
    c1, c2, c3, c4 = st.columns([1.5, 1, 1.5, 1.5])
    
    with c2:
        st.write(" ")
        # 🔄 이 버튼을 누를 때만 캐시를 비우고 새로 읽어옵니다.
        update_trigger = st.button("🔄 상황판 데이터 갱신", use_container_width=True, type="primary")
        if update_trigger:
            st.cache_data.clear() # 기존에 저장된 데이터 삭제
            st.rerun() # 새로고침하여 get_v7_minimal_memo() 재실행

    # 데이터 가져오기
    df_o_raw, df_r_raw = get_v7_minimal_memo()

    # 데이터가 비어있지 않을 때만 계산 시작
    if not df_o_raw.empty:
        q_col = next((c for c in ["추가발주", "추가발주량", "수량"] if c in df_o_raw.columns), df_o_raw.columns[6])
        date_col = next((c for c in ["날짜", "발주시간"] if c in df_o_raw.columns), df_o_raw.columns[0])

        # 데이터 전처리
        for df in [df_o_raw, df_r_raw]:
            if not df.empty:
                df[q_col] = pd.to_numeric(df[q_col], errors='coerce').fillna(0).astype(int)
                df['key'] = (df['상품명'].astype(str) + df['옵션'].astype(str)).str.replace(" ","").str.upper()
                df['short_date'] = pd.to_datetime(df[date_col], errors='coerce').dt.strftime('%m/%d').fillna('')

        # 1. 발주 집계 (이모지 제거 및 중복 제거)
        df_o_raw['memo_clean'] = df_o_raw.apply(
            lambda x: f"{x['short_date']} {x['메모']}" if str(x.get('메모','')).strip() else "", axis=1
        )
        df_orders = df_o_raw[df_o_raw[q_col] > 0].groupby(['key', '업체명', '상품명', '옵션'], as_index=False).agg({
            date_col: 'max', '공급처상품명': 'first', q_col: 'sum',
            'memo_clean': lambda x: " / ".join(dict.fromkeys(filter(None, x.astype(str))))
        }).rename(columns={q_col: '총리오더수량', 'memo_clean': '발주메모'})

        # 2. 입고 집계
        df_r_minus = df_r_raw[df_r_raw[q_col] < 0].copy() if not df_r_raw.empty else pd.DataFrame()
        if not df_r_minus.empty:
            df_r_minus['memo_clean'] = df_r_minus.apply(
                lambda x: f"{x['short_date']} {x[q_col]}{x['메모']}" if str(x.get('메모','')).strip() else f"{x['short_date']} {x[q_col]}개", axis=1
            )
            df_receives = df_r_minus.groupby('key').agg({
                q_col: lambda x: abs(x.sum()),
                'memo_clean': lambda x: " / ".join(dict.fromkeys(filter(None, x.astype(str))))
            }).reset_index().rename(columns={q_col: '입고수량', 'memo_clean': '입고메모'})
        else:
            df_receives = pd.DataFrame(columns=['key', '입고수량', '입고메모'])

        # 3. 통합
        df_total = pd.merge(df_orders, df_receives, on='key', how='left')
        df_total['입고수량'] = df_total['입고수량'].fillna(0).astype(int)
        df_total['미입고잔량'] = df_total['총리오더수량'] - df_total['입고수량']
        
        def combine_memos(row):
            m = []
            if row['발주메모']: m.append(f"[발주] {row['발주메모']}")
            if row['입고메모']: m.append(f"[입고] {row['입고메모']}")
            return " | ".join(m)
            
        df_total['메모이력'] = df_total.apply(combine_memos, axis=1)
        df_total = df_total[df_total['미입고잔량'] > 0].copy()

        # --- [4. 필터 UI (나머지 컬럼들)] ---
        with c1: search_date = st.date_input("📅 날짜 범위", value=[])
        with c3: search_prod = st.text_input("📦 상품 검색", placeholder="상품명/옵션")
        with c4:
            v_list = ["전체 업체"] + sorted(df_total["업체명"].unique().tolist())
            search_vendor = st.selectbox("🏭 업체 선택", v_list)

        # --- [5. 업체별 요약] ---
        st.write("#### 🏭 업체별 미입고 요약")
        v_summary = df_total.groupby("업체명")["미입고잔량"].sum().reset_index().sort_values("미입고잔량", ascending=False)
        if not v_summary.empty:
            m_cols = st.columns(5)
            for i, row in enumerate(v_summary.itertuples()):
                with m_cols[i % 5]:
                    st.metric(label=row.업체명, value=f"{int(row.미입고잔량):,}개")
        st.divider()

        # --- [6. 상세 리스트] ---
        df_disp = df_total.copy()
        if search_vendor != "전체 업체": df_disp = df_disp[df_disp["업체명"] == search_vendor]
        if search_prod: df_disp = df_disp[df_disp["상품명"].str.contains(search_prod, case=False) | df_disp["옵션"].str.contains(search_prod, case=False)]
        
        if not df_disp.empty:
            display_cols = ["날짜", "업체명", "상품명", "옵션", "공급처상품명", "총리오더수량", "입고수량", "미입고잔량", "메모이력"]
            st.dataframe(
                df_disp.sort_values("날짜", ascending=False).rename(columns={date_col: "날짜"}),
                use_container_width=True, hide_index=True,
                column_order=display_cols,
                column_config={
                    "총리오더수량": st.column_config.NumberColumn("총발주"),
                    "입고수량": st.column_config.NumberColumn("입고"),
                    "미입고잔량": st.column_config.NumberColumn("잔량"),
                    "메모이력": st.column_config.TextColumn("메모 (날짜/수량/내용)")
                }
            )
            
            # 엑셀 다운로드
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_disp[display_cols].to_excel(writer, index=False, sheet_name='미입고')
            st.download_button(label="📥 엑셀 다운로드", data=output.getvalue(), file_name="미입고현황.xlsx")
    else:
        st.info("💡 상황판을 갱신하려면 상단의 '상황판 데이터 갱신' 버튼을 눌러주세요.")
