import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from datetime import datetime, timedelta
import re

st.set_page_config(layout="wide", page_title="쿠팡 VF & 평택(본창고) 마스터 발주 대시보드")
st.title("📦 쿠팡 VF & 평택물류 통합 마스터 발주 시스템")
st.caption("쿠팡 PO 계단식 차감 + 평택창고 생산리드타임 역산 실시간 발주 데드라인 연산기")

# 1. 사이드바 설정
st.sidebar.header("📁 엑셀 파일 업로드")
uploaded_file = st.sidebar.file_uploader("통합 마스터 엑셀 업로드 (.xlsx)", type=["xlsx"])

st.sidebar.divider()
st.sidebar.header("⚙️ 발주 및 시뮬레이션 기준")
base_date_input = st.sidebar.date_input("재고/출고 기준일자", datetime(2026, 9, 15).date())
recent_days_window = st.sidebar.number_input(
    "출고 평균 산출 기간 (일)", 
    min_value=1, 
    max_value=31, 
    value=7, 
    help="기준일자 직전 며칠간의 일일 출고량을 평균 낼지 설정합니다. (기본: 최근 7일)"
)
default_lead_time = st.sidebar.number_input("공장 생산/입고 리드타임 (일)", min_value=1, max_value=120, value=35)
vf_trigger_default = st.sidebar.number_input("쿠팡 PO 발생 기준 (VF 잔여량)", min_value=0, value=200, step=10, help="VF 잔여 재고가 이 숫자 밑으로 떨어지면 쿠팡이 평택창고에서 MOQ 단위로 차감해 갑니다.")

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
            if any(re.search(r"^0901$|^901$", v) for v in row_vals) and date_row_idx is None:
                date_row_idx = r

        if sku_row_idx is None:
            sku_row_idx = 1
        if date_row_idx is None:
            date_row_idx = sku_row_idx + 1

        data_start_row = max(sku_row_idx, date_row_idx) + 1

        # 헤더 텍스트 결합
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

        date_col_map = {}
        inbound_col_info = []

        base_year = base_date_input.year

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

            # 일일 출고 컬럼 탐색
            match_date = re.search(r"(\d{2})(\d{2})", h)
            if match_date and ("입고" not in h) and ("재고" not in h):
                m_int = int(match_date.group(1))
                d_int = int(match_date.group(2))
                if 1 <= m_int <= 12 and 1 <= d_int <= 31:
                    try:
                        col_dt = datetime(base_year, m_int, d_int).date()
                        date_col_map[col_dt] = idx
                    except:
                        pass

            # 우측 입고 스케줄 탐색
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

        # 기준일자 직전 N일 날짜 컬럼 추출
        target_calc_dates = []
        for d_back in range(1, int(recent_days_window) + 1):
            check_d = base_date_input - timedelta(days=d_back)
            if check_d in date_col_map:
                target_calc_dates.append((check_d, date_col_map[check_d]))

        if not target_calc_dates and (base_date_input in date_col_map):
            target_calc_dates.append((base_date_input, date_col_map[base_date_input]))

        period_desc = f"{target_calc_dates[-1][0].strftime('%m/%d')} ~ {target_calc_dates[0][0].strftime('%m/%d')} ({len(target_calc_dates)}일간)" if target_calc_dates else "날짜 컬럼 불일치"

        # 데이터 변환
        processed_items = []
        for _, r in df_rows.iterrows():
            item_name = str(r[sku_col_idx]).strip()
            item_cat = str(r[category_col_idx]).strip() if category_col_idx is not None and pd.notna(r[category_col_idx]) else "기타"
            item_moq = clean_num(r[moq_col_idx]) if moq_col_idx is not None else 100.0
            item_vf = clean_num(r[vf_col_idx]) if vf_col_idx is not None else 0.0
            item_main = clean_num(r[main_col_idx]) if main_col_idx is not None else 0.0
            item_safety = clean_num(r[safety_col_idx]) if safety_col_idx is not None else 0.0

            if target_calc_dates:
                sum_recent_sales = sum([clean_num(r[c_idx]) for _, c_idx in target_calc_dates])
                adu = sum_recent_sales / float(len(target_calc_dates))
            else:
                adu = 0.0

            inbounds = []
            for ic_idx, m, d in inbound_col_info:
                qty = clean_num(r[ic_idx])
                if qty > 0:
                    y = base_year if m >= base_date_input.month else base_year + 1
                    in_d = datetime(y, m, d).date()
                    if in_d >= base_date_input:
                        inbounds.append((in_d, int(qty)))

            inbounds.sort(key=lambda x: x[0])

            processed_items.append({
                "구분": item_cat,
                "품명": item_name,
                "평택(본창고)": int(item_main),
                "VF재고": int(item_vf),
                "납품MOQ": int(item_moq) if item_moq > 0 else 100,
                "안전재고": int(item_safety),
                "일출고량": adu,
                "입고스케줄": inbounds
            })

        df_items = pd.DataFrame(processed_items)

        # 2. 카테고리 필터
        st.divider()
        all_cats = ["전체 보기"] + [c for c in df_items["구분"].unique() if c not in ["nan", "", "기타"]] + ["기타"]
        all_cats = list(dict.fromkeys(all_cats))
        selected_cat = st.selectbox("📂 구분(카테고리) 필터:", all_cats)

        view_items = df_items if selected_cat == "전체 보기" else df_items[df_items["구분"] == selected_cat]

        # 3. 전 품목 시뮬레이션 연산
        sim_days = 90
        sim_dates = [base_date_input + timedelta(days=i) for i in range(sim_days)]

        master_table_rows = []

        for _, item in view_items.iterrows():
            c_main = item["평택(본창고)"]
            c_vf = item["VF재고"]
            moq = item["납품MOQ"]
            adu = item["일출고량"]
            safety = item["안전재고"]
            inbounds = item["입고스케줄"]

            has_safety = safety > 0
            limit_target = safety if has_safety else 0

            main_breach_day = None
            first_po_day = None

            for d_i in range(sim_days):
                curr_d = sim_dates[d_i]

                for in_d, in_q in inbounds:
                    if in_d == curr_d:
                        c_main += in_q

                if adu > 0:
                    c_vf -= adu

                if adu > 0 and c_vf < vf_trigger_default:
                    if first_po_day is None:
                        first_po_day = curr_d

                    if c_main >= moq:
                        c_main -= moq
                        c_vf += moq

                        if has_safety and c_main <= limit_target and main_breach_day is None:
                            main_breach_day = d_i
                    else:
                        if main_breach_day is None:
                            main_breach_day = d_i
                        c_vf += c_main
                        c_main = 0

            if adu == 0:
                status = "💤 출고없음 (안전)"
                po_txt = "-"
                breach_txt = "고갈없음"
                deadline_txt = "-"
                d_day_val = 999
            elif main_breach_day is None:
                status = "✅ [안정] 90일 이상 여유"
                po_txt = first_po_day.strftime('%m/%d') if first_po_day else "-"
                breach_txt = "90일 이상 버팀"
                deadline_txt = "여유"
                d_day_val = 999
            else:
                b_date = sim_dates[main_breach_day]
                breach_txt = f"{b_date.strftime('%m/%d')} (D+{main_breach_day}일)"
                order_deadline = b_date - timedelta(days=default_lead_time)
                deadline_txt = order_deadline.strftime('%Y-%m-%d')
                d_day_val = (order_deadline - base_date_input).days
                po_txt = first_po_day.strftime('%m/%d') if first_po_day else "-"

                if d_day_val <= 0:
                    status = "🚨 [초긴급] 지금 즉시 발주!"
                elif d_day_val <= 7:
                    status = f"⚠️ [긴급] D-{d_day_val}일 내 발주"
                elif d_day_val <= 14:
                    status = f"🔔 [준비] D-{d_day_val}일 발주준비"
                else:
                    status = f"✅ [안정] D-{d_day_val}일 여유"

            master_table_rows.append({
                "구분": item["구분"],
                "품명": item["품명"],
                "평택(본창고)": f"{item['평택(본창고)']:,}개",
                "VF재고": f"{item['VF재고']:,}개",
                "납품MOQ": f"{moq:,}개",
                "안전재고": f"{safety:,}개" if has_safety else "미적용",
                "일일판매량": f"{adu:.1f}개/일",
                "차기 쿠팡PO예정일": po_txt,
                "평택재고 고갈일": breach_txt,
                "공장발주 데드라인": deadline_txt,
                "발주상태": status,
                "raw_name": item["품명"],
                "d_day_num": d_day_val
            })

        summary_df = pd.DataFrame(master_table_rows)

        # 발주상태 우선순위 정렬 후, 번호(인덱스)를 1, 2, 3... 순으로 깔끔하게 재설정
        sorted_table = summary_df.drop(columns=["raw_name", "d_day_num"]).sort_values(by="발주상태").reset_index(drop=True)
        sorted_table.index = sorted_table.index + 1  # 1번부터 시작하도록 설정

        # KPI
        u_cnt = len(summary_df[summary_df["발주상태"].str.contains("초긴급")])
        w_cnt = len(summary_df[summary_df["발주상태"].str.contains("긴급|준비")])

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("📅 출고 평균 산출 구간", period_desc)
        k2.metric("🚨 즉시 공장 발주 필요", f"{u_cnt} 개 SKU", delta_color="inverse")
        k3.metric("⚠️ 14일 내 발주 예정", f"{w_cnt} 개 SKU")
        k4.metric("🏭 적용 생산 리드타임", f"{default_lead_time}일 소요")

        st.subheader("📋 전체 품목 쿠팡 PO 연동 및 평택창고 발주 데드라인 마스터 테이블")
        st.dataframe(sorted_table, use_container_width=True)

        # 4. 하단 개별 정밀 시뮬레이터 & 그래프
        st.divider()
        st.subheader("🔍 개별 SKU 쿠팡 PO 차감 타임라인 & 재고 곡선")

        selected_sku = st.selectbox("정밀 조회할 품목을 선택하세요:", df_items["품명"].unique())
        target_item = df_items[df_items["품명"] == selected_sku].iloc[0]

        s1, s2, s3, s4 = st.columns(4)
        with s1:
            p_trig = st.number_input("쿠팡 PO 트리거 (VF 잔여)", min_value=0, value=int(vf_trigger_default), step=10)
        with s2:
            p_mult = st.selectbox("쿠팡 PO 배수", [1, 2, 3], format_func=lambda x: f"{x}배수 ({target_item['납품MOQ']*x:,}개)")
            active_po_unit = target_item['납품MOQ'] * p_mult
        with s3:
            p_adu = st.number_input("일일 고객 출고량 (개/일)", min_value=0.0, value=float(target_item["일출고량"]), step=1.0)
        with s4:
            p_safety = st.number_input("안전재고 수량 (0이면 품절기준)", min_value=0, value=int(target_item["안전재고"]), step=100)

        c_main = target_item["평택(본창고)"]
        c_vf = target_item["VF재고"]
        inbound_sched = target_item["입고스케줄"]

        hist_main = []
        hist_vf = []
        event_logs = []
        ind_breach = None
        ind_first_po = None

        for d_i in range(sim_days):
            curr_d = sim_dates[d_i]

            for in_d, in_q in inbound_sched:
                if in_d == curr_d:
                    c_main += in_q
                    event_logs.append((curr_d, f"🚚 [평택 입고] +{in_q:,}개 공장 입고 완료", c_main, c_vf))

            if p_adu > 0:
                c_vf -= p_adu

            if p_adu > 0 and c_vf < p_trig:
                if ind_first_po is None:
                    ind_first_po = curr_d

                if c_main >= active_po_unit:
                    c_main -= active_po_unit
                    c_vf += active_po_unit

                    if p_safety > 0 and c_main <= p_safety and ind_breach is None:
                        ind_breach = d_i

                    event_logs.append((curr_d, f"📦 쿠팡 PO 차감: -{active_po_unit:,}개 ➔ 평택잔여: {c_main:,}개", c_main, c_vf))
                else:
                    if ind_breach is None:
                        ind_breach = d_i
                    event_logs.append((curr_d, f"🚨 [평택 결품] PO {active_po_unit:,}개 중 잔여 {c_main:,}개만 VF 이동", 0, c_vf + c_main))
                    c_vf += c_main
                    c_main = 0

            hist_main.append(c_main)
            hist_vf.append(max(c_vf, 0))

        st.markdown(f"#### 📈 [{selected_sku}] 향후 60일 평택본창고 vs VF 재고 흐름")
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=sim_dates[:60], y=hist_main[:60], mode="lines+markers", name="평택본창고 (쿠팡 PO 시 계단식 차감)", line=dict(color="#1f77b4", width=3)))
        fig.add_trace(go.Scatter(x=sim_dates[:60], y=hist_vf[:60], mode="lines", name="VF 잔여재고 (고객 일일출고로 차감)", line=dict(color="#ff7f0e", width=2, dash="dash")))

        if p_safety > 0:
            fig.add_hline(y=p_safety, line_dash="dashdot", line_color="#ffbb00", annotation_text=f"안전재고 마지노선 ({p_safety:,}개)")
        fig.add_hline(y=p_trig, line_dash="dot", line_color="red", annotation_text=f"쿠팡 발주 트리거 ({p_trig}개)")

        fig.update_layout(xaxis_title="일자", yaxis_title="수량 (개)", hovermode="x unified", margin=dict(l=20, r=20, t=30, b=20))
        st.plotly_chart(fig, use_container_width=True)

        with st.expander("📅 일자별 쿠팡 PO 발생 및 평택창고 이동 타임라인"):
            if event_logs:
                st.dataframe(pd.DataFrame(event_logs, columns=["일자", "내용", "평택 잔여", "VF 잔여"]), use_container_width=True)

    except Exception as e:
        st.error(f"오류가 발생했습니다: {e}")
else:
    st.info("👈 왼쪽 사이드바에서 마스터 엑셀 파일을 업로드해 주세요.")
