import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from datetime import datetime, timedelta
import re

st.set_page_config(layout="wide", page_title="쿠팡 VF & 평택(본창고) 마스터 발주 대시보드")
st.title("📦 쿠팡 VF & 평택물류 통합 마스터 발주 시스템")
st.caption("일일 출고량(최신 7일 자동 추적) + 쿠팡 납품 MOQ + 평택/VF 실재고 + 안전재고 + 우측 입고 스케줄 통합 연동")

# 1. 사이드바 설정
st.sidebar.header("📁 엑셀 파일 업로드")
uploaded_file = st.sidebar.file_uploader("통합 마스터 엑셀 업로드 (.xlsx)", type=["xlsx"])

st.sidebar.divider()
st.sidebar.header("⚙️ 분석 및 발주 기준 설정")
base_date_input = st.sidebar.date_input("재고 기준일자", datetime(2026, 9, 15).date())
recent_days_window = st.sidebar.number_input(
    "최근 출고량 반영 기간 (일)", 
    min_value=1, 
    max_value=31, 
    value=7, 
    help="가장 최근 출고가 발생한 날짜로부터 직전 며칠간의 평균을 일일출고량으로 볼 것인지 설정합니다. (기본: 최근 7일)"
)
default_lead_time = st.sidebar.number_input("기본 구매 리드타임 (일)", min_value=1, max_value=120, value=35)
vf_trigger_default = st.sidebar.number_input("기본 쿠팡 PO 트리거 (VF 잔여)", min_value=0, value=200, step=10)

def clean_txt(val):
    if pd.isna(val):
        return ""
    return str(val).strip().replace("\n", "").replace(" ", "")

def clean_num(val):
    if pd.isna(val) or str(val).strip() in ["-", "", "#DIV/0!", "전량", "nan", "NaN", "O", "X"]:
        return 0.0
    try:
        return float(str(val).replace(",", "").strip())
    except:
        return 0.0

if uploaded_file:
    try:
        xl = pd.ExcelFile(uploaded_file)
        sheet_name = st.sidebar.selectbox("조회할 시트 선택:", xl.sheet_names, index=0)

        raw_df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)

        sku_row_idx = None
        date_row_idx = None

        for r in range(min(6, len(raw_df))):
            row_vals = [clean_txt(x) for x in raw_df.iloc[r].values]
            row_str = " ".join(row_vals)
            if "품명" in row_str and sku_row_idx is None:
                sku_row_idx = r
            if any(re.search(r"0901|901", v) for v in row_vals) and date_row_idx is None:
                date_row_idx = r

        if sku_row_idx is None:
            sku_row_idx = 1
        if date_row_idx is None:
            date_row_idx = sku_row_idx + 1

        data_start_row = max(sku_row_idx, date_row_idx) + 1

        col_headers = []
        for c in range(raw_df.shape[1]):
            pieces = []
            for r in range(data_start_row):
                val = clean_txt(raw_df.iloc[r, c])
                if val and val not in pieces:
                    pieces.append(val)
            col_headers.append("_".join(pieces))

        sku_col_idx = None
        category_col_idx = None
        moq_col_idx = None
        vf_col_idx = None
        main_col_idx = None
        safety_col_idx = None

        date_col_indices = []
        inbound_col_info = []

        for idx, h in enumerate(col_headers):
            h_upper = h.upper()

            if "품명" in h and sku_col_idx is None:
                sku_col_idx = idx
            elif ("구분" in h or "유형" in h) and category_col_idx is None:
                category_col_idx = idx
            elif ("MOQ" in h_upper or "납품" in h) and moq_col_idx is None:
                moq_col_idx = idx
            elif "안전재고" in h and safety_col_idx is None:
                safety_col_idx = idx
            elif ("VF" in h_upper and ("재고" in h or "수량" in h)) and vf_col_idx is None:
                vf_col_idx = idx
            elif (("평택" in h or "본창고" in h or "본물류" in h) and "재고" in h and "안전" not in h) and main_col_idx is None:
                main_col_idx = idx

            # 일자별 출고 컬럼 (0901~0931, 1001~1031 등 4자리 날짜 패턴)
            match_date = re.search(r"(\d{2})(\d{2})", h)
            if match_date and ("입고" not in h) and ("재고" not in h):
                date_col_indices.append((idx, h, match_date.group(0)))

            # 우측 입고일정 컬럼 매칭
            if "입고" in h:
                m_match = re.search(r"(\d{1,2})월\s*(\d{1,2})일", h)
                if m_match:
                    inbound_col_info.append((idx, int(m_match.group(1)), int(m_match.group(2))))

        if sku_col_idx is None:
            sku_col_idx = 1

        df_rows = raw_df.iloc[data_start_row:].copy()
        valid_indices = []
        for r_i, r_val in df_rows.iterrows():
            name = str(r_val[sku_col_idx]).strip()
            if pd.isna(r_val[sku_col_idx]) or name in ["", "nan", "NaN", "합계", "평균", "비고"]:
                continue
            if "합계" in name:
                continue
            valid_indices.append(r_i)

        df_rows = df_rows.loc[valid_indices]

        # 🌟 핵심: 실제 출고 데이터가 입력되어 있는 '가장 최신 날짜 컬럼' 자동 탐색
        active_date_cols = []
        for c_idx, h_name, d_str in date_col_indices:
            # 전체 품목 합산 출고량이 0보다 큰 컬럼만 유효 출고일자로 판정
            col_total = df_rows[c_idx].apply(clean_num).sum()
            if col_total > 0:
                active_date_cols.append((c_idx, h_name, d_str))

        # 최근 N일(기본 7일) 슬라이싱
        if active_date_cols:
            selected_recent_dates = active_date_cols[-int(recent_days_window):]
            recent_col_indices = [c[0] for c in selected_recent_dates]
            period_label = f"{selected_recent_dates[0][2]} ~ {selected_recent_dates[-1][2]} (최근 {len(selected_recent_dates)}일 평균)"
        else:
            recent_col_indices = []
            period_label = "출고 내역 없음"

        # 데이터 변환
        processed_data = []
        for _, r in df_rows.iterrows():
            item_name = str(r[sku_col_idx]).strip()
            item_cat = str(r[category_col_idx]).strip() if category_col_idx is not None and pd.notna(r[category_col_idx]) else "기타"
            item_moq = clean_num(r[moq_col_idx]) if moq_col_idx is not None else 100.0
            item_vf = clean_num(r[vf_col_idx]) if vf_col_idx is not None else 0.0
            item_main = clean_num(r[main_col_idx]) if main_col_idx is not None else 0.0
            item_safety = clean_num(r[safety_col_idx]) if safety_col_idx is not None else 0.0

            # 최근 N일치 합산 후 일수(7일)로 나누어 최신 ADU 계산
            if recent_col_indices:
                recent_sum = sum([clean_num(r[c_i]) for c_i in recent_col_indices])
                adu = recent_sum / float(len(recent_col_indices))
            else:
                adu = 0.0

            inbounds = []
            for ic_idx, m, d in inbound_col_info:
                qty = clean_num(r[ic_idx])
                if qty > 0:
                    y = base_date_input.year if m >= base_date_input.month else base_date_input.year + 1
                    inbounds.append((datetime(y, m, d).date(), int(qty)))

            processed_data.append({
                "구분": item_cat,
                "품명": item_name,
                "평택(본창고)": int(item_main),
                "VF재고": int(item_vf),
                "납품MOQ": int(item_moq) if item_moq > 0 else 100,
                "안전재고": int(item_safety),
                "일평균출고": adu,
                "입고스케줄": inbounds
            })

        df_main = pd.DataFrame(processed_data)

        # 2. 카테고리 필터
        st.divider()
        all_cats = ["전체 보기"] + [c for c in df_main["구분"].unique() if c not in ["nan", "", "기타"]] + ["기타"]
        all_cats = list(dict.fromkeys(all_cats))
        selected_cat = st.selectbox("📂 구분(카테고리) 선택:", all_cats)

        if selected_cat != "전체 보기":
            view_data = df_main[df_main["구분"] == selected_cat].copy()
        else:
            view_data = df_main.copy()

        calc_rows = []
        for _, row in view_data.iterrows():
            tot_stk = row["평택(본창고)"] + row["VF재고"]
            safety = row["안전재고"]
            adu = row["일평균출고"]
            moq = row["납품MOQ"]

            has_safety = safety > 0
            target_limit = safety if has_safety else 0

            days_to_target = max((tot_stk - target_limit) / adu, 0.0) if adu > 0 else 999.0
            d_day = days_to_target - default_lead_time

            if tot_stk <= target_limit and adu > 0:
                status = "🚨 [초긴급] 마지노선 붕괴 (즉시발주)"
            elif d_day <= 0 and adu > 0:
                status = "🔥 [긴급] 오늘 발주 필요"
            elif d_day <= 7 and adu > 0:
                status = f"⚠️ [주의] D-{int(d_day)}일 내 발주"
            elif adu == 0:
                status = "💤 출고없음 (유지)"
            else:
                status = f"✅ [안정] D-{int(d_day)}일 여유"

            calc_rows.append({
                "구분": row["구분"],
                "품명": row["품명"],
                "평택(본창고)": f"{row['평택(본창고)']:,}개",
                "VF재고": f"{row['VF재고']:,}개",
                "납품MOQ": f"{moq:,}개",
                "안전재고": f"{safety:,}개" if has_safety else "미적용",
                "최근일출고": f"{adu:.1f}개/일",
                "목표도달소진일": f"{days_to_target:.1f}일" if days_to_target < 900 else "-",
                "발주상태": status,
                "raw_name": row["품명"],
                "d_day_val": d_day
            })

        table_df = pd.DataFrame(calc_rows)

        u_cnt = len(table_df[table_df["발주상태"].str.contains("초긴급|오늘")])
        w_cnt = len(table_df[table_df["발주상태"].str.contains("주의")])

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("재고 기준일자", base_date_input.strftime("%Y-%m-%d"))
        k2.metric("📊 출고 계산 구간", period_label)
        k3.metric("🚨 즉시 발주 대상", f"{u_cnt} 개 SKU")
        k4.metric("⚠️ 이번 주 발주 검토", f"{w_cnt} 개 SKU")

        st.subheader("📋 전체 품목 발주 현황 마스터 테이블")
        st.dataframe(
            table_df.drop(columns=["raw_name", "d_day_val"]).sort_values(by="발주상태"),
            use_container_width=True
        )

        # 3. 개별 시뮬레이션
        st.divider()
        st.subheader("🔍 개별 SKU 정밀 시뮬레이터 (우측 입고 스케줄 자동 연동)")

        selected_sku = st.selectbox("정밀 분석할 품목을 선택하세요:", df_main["품명"].unique())
        item = df_main[df_main["품명"] == selected_sku].iloc[0]

        s1, s2, s3, s4 = st.columns(4)
        with s1:
            sim_vf_trig = st.number_input("쿠팡 발주 트리거 (VF 잔여)", min_value=0, value=int(vf_trigger_default), step=10)
        with s2:
            sim_mult = st.selectbox("쿠팡 PO 배수", [1, 2, 3], format_func=lambda x: f"{x}배수 ({item['납품MOQ']*x:,}개)")
            po_unit = item['납품MOQ'] * sim_mult
        with s3:
            sim_adu = st.number_input("최근 일평균 출고량 (개/일)", min_value=0.0, value=float(item["일평균출고"]), step=1.0)
        with s4:
            sim_safety = st.number_input("안전재고 수량 (0이면 원래대로)", min_value=0, value=int(item["안전재고"]), step=100)

        sim_days = 90
        sim_dates = [base_date_input + timedelta(days=i) for i in range(sim_days)]

        curr_main = item["평택(본창고)"]
        curr_vf = item["VF재고"]
        inbound_sched = item["입고스케줄"]

        hist_main = []
        hist_vf = []
        po_logs = []

        breach_day = None
        first_po_date = None

        for day_idx in range(sim_days):
            today_d = sim_dates[day_idx]

            for in_d, in_q in inbound_sched:
                if in_d == today_d:
                    curr_main += in_q
                    po_logs.append((today_d, f"🚚 [우측 입고반영] +{in_q:,}개 평택본창고 입고", curr_main, curr_vf))

            if sim_adu > 0:
                curr_vf -= sim_adu

            if sim_adu > 0 and curr_vf < sim_vf_trig:
                if first_po_date is None:
                    first_po_date = today_d

                if curr_main >= po_unit:
                    curr_main -= po_unit
                    curr_vf += po_unit

                    if sim_safety > 0 and curr_main <= sim_safety and breach_day is None:
                        breach_day = day_idx

                    po_logs.append((today_d, f"📦 쿠팡 PO 차감: -{po_unit:,}개 (평택 잔여: {curr_main:,}개)", curr_main, curr_vf))
                else:
                    if breach_day is None:
                        breach_day = day_idx
                    po_logs.append((today_d, f"🚨 [평택 결품] PO {po_unit:,}개 중 잔여 {curr_main:,}개만 이동", 0, curr_vf + curr_main))
                    curr_vf += curr_main
                    curr_main = 0

            hist_main.append(curr_main)
            hist_vf.append(max(curr_vf, 0))

        is_safe_mode = sim_safety > 0
        mode_str = "안전재고 붕괴일" if is_safe_mode else "평택창고 품절일"

        if sim_adu == 0:
            breach_txt = "출고 없음"
            order_txt = "발주 불필요"
            d_day_int = 999
        elif breach_day is None:
            breach_txt = "90일 이상 유지"
            order_txt = "여유"
            d_day_int = 999
        else:
            b_dt = sim_dates[breach_day]
            breach_txt = f"{b_dt.strftime('%Y-%m-%d')} (D+{breach_day}일)"
            o_dt = b_dt - timedelta(days=default_lead_time)
            order_txt = o_dt.strftime("%Y-%m-%d")
            d_day_int = (o_dt - base_date_input).days

        st.markdown("---")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("시작 재고 (09/15)", f"평택 {item['평택(본창고)']:,}개 / VF {item['VF재고']:,}개", f"안전재고: {sim_safety:,}개" if is_safe_mode else "미적용")
        c2.metric("다음 쿠팡 PO 예상", first_po_date.strftime('%m/%d') if first_po_date else "-", f"1회 PO단위: {po_unit:,}개")
        c3.metric(mode_str, breach_txt, f"입고스케줄 {len(inbound_sched)}건 반영됨")
        c4.metric("구매팀 발주 데드라인", order_txt, f"D-Day {d_day_int}일" if d_day_int < 900 else None, delta_color="inverse")

        st.markdown(f"#### 📈 [{selected_sku}] 향후 60일 재고 시뮬레이션")
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=sim_dates[:60], y=hist_main[:60], mode="lines+markers", name="평택본창고 (계단식 차감 & 입고 충전)", line=dict(color="#1f77b4", width=3)))
        fig.add_trace(go.Scatter(x=sim_dates[:60], y=hist_vf[:60], mode="lines", name="VF 잔여 재고", line=dict(color="#ff7f0e", width=2, dash="dash")))

        if is_safe_mode:
            fig.add_hline(y=sim_safety, line_dash="dashdot", line_color="#ffbb00", annotation_text=f"안전재고 마지노선 ({sim_safety:,}개)")
        fig.add_hline(y=sim_vf_trig, line_dash="dot", line_color="red", annotation_text=f"쿠팡 트리거 ({sim_vf_trig}개)")

        fig.update_layout(xaxis_title="날짜", yaxis_title="수량 (개)", hovermode="x unified", margin=dict(l=20, r=20, t=30, b=20))
        st.plotly_chart(fig, use_container_width=True)

        with st.expander("📅 일자별 상세 이벤트 로그"):
            if po_logs:
                st.dataframe(pd.DataFrame(po_logs, columns=["일자", "내용", "평택 잔여", "VF 잔여"]), use_container_width=True)

    except Exception as e:
        st.error(f"오류가 발생했습니다: {e}")
else:
    st.info("👈 왼쪽 사이드바에서 마스터 엑셀 파일을 업로드해 주세요.")
