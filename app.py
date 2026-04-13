File "/mount/src/order-system/app.py", line 161
          st.divider()
         ^
IndentationError: unexpected indent

        # --- [6단계: 히스토리 (최근 저장 내역)] ---
        st.divider()
        st.header("6️⃣ 최근 히스토리 (History)")
        sh_hist = get_sheet()
        if sh_hist:
            ws_h = sh_hist.worksheet("history")
            # 최근 10개 행만 가져오기
            hist_data = ws_h.get_all_values()
            if len(hist_data) > 1:
                h_df = pd.DataFrame(hist_data[1:], columns=hist_data[0]).tail(10)
                st.table(h_df.iloc[::-1]) # 최신순 정렬해서 표로 표시
            else:
                st.write("기록된 히스토리가 없습니다.")

        # --- [7단계: 리오더 현황판 (실시간)] ---
        st.divider()
        st.header("7️⃣ 실시간 리오더 현황판")
        if 'r_map' in st.session_state:
            # 잔량이 있는 항목만 모아서 요약
            summary = []
            for k, v in st.session_state.r_map.items():
                if v > 0: summary.append({"상품 식별키": k, "현재 리오더 총 잔량": v})
            
            if summary:
                st.dataframe(pd.DataFrame(summary), use_container_width=True)
            else:
                st.info("현재 진행 중인 리오더 잔량이 없습니다.")
