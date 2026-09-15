import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from datetime import datetime, timedelta
import re

st.set_page_config(layout="wide", page_title="쿠팡 VF & 평택(본창고) 마스터 발주 대시보드")
st.title("📦 쿠팡 VF & 평택물류 통합 마스터 발주 시스템")
st.caption("일일 출고량(0901~) + 쿠팡 납품 MOQ + 평택/VF 실재고 + 안전재고 + 우측 입고 스케줄 통합 연동")

# 1. 사이드바 설정
st.sidebar.header("📁 엑셀 파일 업로드")
uploaded_file = st.sidebar.file_uploader("통합 마스터 엑셀 업로드 (.xlsx)", type=["xlsx"])

st.sidebar.divider()
st.sidebar.header("⚙️ 기본 설정")
base_date_input = st.sidebar.date_input("재고 기준일자", datetime(2026, 9, 15).date())
default_lead_time = st.sidebar.number_input("기본 구매 리드타임 (일)", min_value=1, max_value=120, value=35)
vf_trigger_default = st.sidebar.number_input("기본 쿠팡 PO 트리거 (VF 잔여)", min_value=0, value=200, step=10)

def clean_txt(val):
    if pd.isna(val):
        return ""
    return str(val).strip().replace("\n", "").replace(" ", "")

def clean_num(val):
    if pd.isna(val) or str(val).strip() in ["-", "", "#DIV/0!", "전량", "nan", "NaN"]:
        return 0.0
    try:
        return float(str(val).replace(",", "").strip())
    except:
        return 0.0

if uploaded_file:
    try:
        xl = pd.ExcelFile(uploaded_file)
        sheet_name = st.sidebar.selectbox("조회할 시트 선택:", xl.sheet_names, index=0)

        # 헤더 없이 원시 데이터 로드
        raw_df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)

        # 엑셀 상단 헤더 행 자동 감지 (품명 또는 0901이 있는 행 탐색)
        header_row_main = 2 # 기본값 3행 (index 2)
        for r in range(min(6, len(raw_df))):
            row_str = " ".join([clean_txt(x) for x in raw_df.iloc[r].values])
            if "0901" in row_str or "품명" in row_str:
                header_row_main = r
                break

        row_1 = raw_df.iloc[header_row_main - 1] if header_row_main > 0 else pd.Series([""] * raw_df.shape[1])
        row_2 = raw_df.iloc[header_row_main]

        # 데이터 시작 행
        df_data = raw_df.iloc[header_row_main + 1:].copy()

        # 컬럼 인덱스 매핑 찾기
        sku_col_idx = None
        category_col_idx = None
        moq_col_idx = None
        vf_col_idx = None
        main_col_idx = None
        safety_col_idx = None

        date_col_indices = []
        inbound_col_info = [] # (col_idx, 월, 일, 컬럼표시명)

        for idx in range(raw_df.shape[1]):
            t1 = clean_txt(row_1[idx])
            t2 = clean_txt(row_2[idx])
            combined = f"{t1}_{t2}"

            # 품명
            if ("품명" in t2 or "품명" in t1) and sku_col_idx is None:
                sku_col_idx = idx
            # 구분/유형
            elif ("구분" in t2 or "구분" in t1 or "유형" in t2) and category_col_idx is None:
                category_col_idx = idx
            # 납품 MOQ
            elif ("MOQ" in combined.upper() or "납품" in combined) and moq_col_idx is None:
                moq_col_idx = idx
            # VF 재고
            elif ("VF" in combined.upper() and ("재고" in combined or "수량" in combined)) and vf_col_idx is None:
                vf_col_idx = idx
            # 평택/본창고 재고
            elif (("평택" in combined or "본창고" in combined or "본물류" in combined) and "재고" in combined) and main_col_idx is None:
                main_col_idx = idx
            # 안전재고
            elif "안전재고" in combined and safety_col_idx is None:
                safety_col_idx = idx

            # 일자별 출고 컬럼 (0901 ~ 0931 등 4자리 월일 패턴 매칭)
            # t2(3행)에 0901, 0902 등이 적혀있음
            if re.search(r"^\d{4}$", t2) or re.search(r"^09\d{2}$", t2):
                date_col_indices.append((idx, t2))

            # 우측 입고일정 컬럼 매칭 (발주/입고일: M월 D일)
            full_header = f"{t1} {t2}"
            if "입고" in full_header:
                m_match = re.search(r"(\d{1,2})월\s*(\d{1,2})일", full_header)
                if m_match:
                    inbound_col_info.append((idx, int(m_match.group(1)), int(m_match.group(2)), f"{m_match.group(1)}/{m_match.group(2)}입고"))

        # 품명 인덱스 예외 보정
        if sku_col_idx is None:
            sku_col_idx = 1

        # 품명이 없거나 합계/평균 행 필터링
        valid_rows = []
        for r_i, r_val in df_data.iterrows():
            name = str(r_val[sku_col_idx]).strip()
            if pd.isna(r_val[sku_col_idx]) or name in ["", "nan", "NaN", "합계", "평균", "비고"]:
                continue
            if "합계" in name:
                continue
            valid_rows.append(r_i)

        df_data = df_data.loc[valid_rows].copy()

        # 데이터 변환 및 정리
        processed_data = []
        for _, r in df_data.iterrows():
            item_name = str(r[sku_col_idx]).strip()
            item_cat = str(r[category_col_idx]).strip() if category_col_idx is not None and pd.notna(r[category_col_idx]) else "기타"
            item_moq = clean_num(r[moq_col_idx]) if moq_col_idx is not None else 100.0
            item_vf = clean_num(r[vf_col_idx]) if vf_col_idx is not None else 0.0
            item_main = clean_num(r[main_col_idx]) if main_col_idx is not None else 0.0
            item_safety = clean_num(r[safety_col_idx]) if safety_col_idx is not None else 0.0

            # 일일 출고량 계산 (앞쪽 0901~0915 컬럼들 중 값이 있는 일자들로 평균)
            out_vals = [clean_num(r[d_idx]) for d_idx, _ in date_col_indices]
            # 0보다 큰 출고가 있거나 기록된 날짜 기준 평균 (0901~0915 등)
            valid_outs = [v for v in out_vals if v > 0]
            # 출고량이 기재된 날짜 수가 있으면 그 일수로 나누고, 없으면 15일치(0901~0915) 기준 평균
            adu = sum(out_vals) / 15.0 if len(out_vals) >= 15 else (np.mean(out_vals) if out_vals else 0.0)

            # 우측 입고 스케줄 파싱
            inbounds = []
            for ic_idx, m, d, label in inbound_col_info:
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
        all_cats = ["전체 보기"] + [c for c in df_main["구분"].unique() if c != "nan" and c != ""]
        selected_cat = st.selectbox("📂 구분(카테고리) 선택:", all_cats)

        if selected_cat != "전체 보기":
            view_data = df_main[df_main["구분"] == selected_cat].copy()
        else:
            view_data = df_main.copy()

        # 전체 테이블 지표 계산
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
                "일출고량": f"{adu:.1f}개/일",
                "목표도달소진일": f"{days_to_target:.1f}일" if days_to_target < 900 else "-",
                "발주상태": status,
                "raw_name": row["품명"],
                "d_day_val": d_day
            })

        table_df = pd.DataFrame(calc_rows)

        # KPI
        u_cnt = len(table_df[table_df["발주상태"].str.contains("초긴급|오늘")])
        w_cnt = len(table_df[table_df["발주상태"].str.contains("주의")])

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("기준 일자", base_date_input.strftime("%Y-%m-%d"))
        k2.metric("🚨 즉시 발주 대상", f"{u_cnt} 개 SKU")
        k3.metric("⚠️ 이번 주 발주 검토", f"{w_cnt} 개 SKU")
        k4.metric("📦 전체 품목 수", f"{len(table_df)} 개")

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
            sim_adu = st.number_input("일평균 출고량 (개/일)", min_value=0.0, value=float(item["일평균출고"]), step=1.0)
        with s4:
            sim_safety = st.number_input("안전재고 수량 (0이면 원래대로)", min_value=0, value=int(item["안전재고"]), step=100)

        # 90일 시뮬레이션
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

            # 우측 입고 자동 반영
            for in_d, in_q in inbound_sched:
                if in_d == today_d:
                    curr_main += in_q
                    po_logs.append((today_d, f"🚚 [우측 입고반영] +{in_q:,}개 평택본창고 입고", curr_main, curr_vf))

            # 출고
            if sim_adu > 0:
                curr_vf -= sim_adu

            # 쿠팡 PO 발생
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

        # 카드 계산
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
        c1.metric("시작 재고", f"평택 {item['평택(본창고)']:,}개 / VF {item['VF재고']:,}개", f"안전재고: {sim_safety:,}개" if is_safe_mode else "미적용")
        c2.metric("다음 쿠팡 PO 예상", first_po_date.strftime('%m/%d') if first_po_date else "-", f"1회 PO단위: {po_unit:,}개")
        c3.metric(mode_str, breach_txt, f"입고스케줄 {len(inbound_sched)}건 반영됨")
        c4.metric("구매팀 발주 데드라인", order_txt, f"D-Day {d_day_int}일" if d_day_int < 900 else None, delta_color="inverse")

        # 그래프
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
