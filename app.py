from datetime import datetime, timedelta
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(layout="wide", page_title="쿠팡 VF & 본창고 통합 발주 관리기")
st.title("📦 쿠팡 VF & 본창고 유형별 구매발주 최적화 대시보드")
st.caption(
    "엑셀 업로드 시 열 이름 자동 매칭 및 쿠팡 PO/본창고 계단식 차감 정밀 시뮬레이션을 제공합니다."
)

# 1. 사이드바 - 파일 업로드 및 기본 설정
st.sidebar.header("📁 엑셀 파일 업로드")
uploaded_file = st.sidebar.file_uploader(
    "통합 관리 엑셀 파일 (.xlsx)", type=["xlsx"]
)

st.sidebar.divider()
st.sidebar.header("⚙️ 구매 기준 설정")
lead_time = st.sidebar.number_input(
    "기본 구매 리드타임 (일)", min_value=1, max_value=90, value=35
)
safety_days = st.sidebar.number_input(
    "안전재고 버퍼 (일)", min_value=0, max_value=30, value=7
)


# 스마트 열 탐지 헬퍼 함수
def find_best_column(columns, candidates):
  # 1. 완전 일치 또는 강력 일치 우선
  for cand in candidates:
    for i, col in enumerate(columns):
      col_clean = str(col).strip().replace(" ", "").upper()
      cand_clean = cand.strip().replace(" ", "").upper()
      if (
          col_clean == cand_clean
          or col_clean.startswith(cand_clean)
          or cand_clean in col_clean
      ):
        return i
  return 0


if uploaded_file:
  try:
    excel_file = pd.ExcelFile(uploaded_file)
    sheet_names = excel_file.sheet_names

    st.sidebar.divider()
    st.sidebar.header("📑 시트 설정")

    # 시트 자동 탐지
    summary_candidates = [
        s for s in sheet_names if any(k in s for k in ["요약", "재고", "현황"])
    ]
    summary_idx = (
        sheet_names.index(summary_candidates[0]) if summary_candidates else 0
    )
    summary_sheet = st.sidebar.selectbox(
        "1. 요약 시트:", sheet_names, index=summary_idx
    )

    out_candidates = [
        s for s in sheet_names if any(k in s for k in ["출고", "일일", "고객"])
    ]
    out_idx = (
        sheet_names.index(out_candidates[0])
        if out_candidates
        else (1 if len(sheet_names) > 1 else 0)
    )
    outbound_sheet = st.sidebar.selectbox(
        "2. 일일 출고량 시트:", sheet_names, index=out_idx
    )

    # 엑셀 데이터 로드 (여백 없앤 1행 기준 기본 로드)
    df_summary = pd.read_excel(uploaded_file, sheet_name=summary_sheet)
    df_daily_out = pd.read_excel(uploaded_file, sheet_name=outbound_sheet)

    # 헤더가 2행에 남아있는 경우 대비 자동 헤더 보정
    if any(str(c).startswith("Unnamed") for c in df_summary.columns[:3]):
      df_summary = pd.read_excel(
          uploaded_file, sheet_name=summary_sheet, header=1
      )
    if any(str(c).startswith("Unnamed") for c in df_daily_out.columns[:3]):
      df_daily_out = pd.read_excel(
          uploaded_file, sheet_name=outbound_sheet, header=1
      )

    df_summary.columns = [str(c).strip() for c in df_summary.columns]
    df_daily_out.columns = [str(c).strip() for c in df_daily_out.columns]

    df_summary = df_summary.dropna(how="all")
    df_daily_out = df_daily_out.dropna(how="all")

    st.sidebar.divider()
    st.sidebar.header("🏷️ 열 매칭 (자동 감지 완료)")

    # 1) 요약 시트 열 자동 매칭
    sku_idx = find_best_column(
        df_summary.columns, ["품명", "상품명", "SKU", "품목", "상품", "품목명"]
    )
    sku_col = st.sidebar.selectbox(
        "[요약] 품명/SKU 열:", df_summary.columns, index=sku_idx
    )

    main_idx = find_best_column(
        df_summary.columns,
        ["본물류", "본창고", "본물류재고", "본창고재고", "물류재고", "본", "창고"],
    )
    main_col = st.sidebar.selectbox(
        "[요약] 본창고 재고 열:", df_summary.columns, index=main_idx
    )

    vf_idx = find_best_column(
        df_summary.columns,
        ["VF잔여", "VF재고", "VF", "플렉스", "벤더플렉스", "쿠팡재고"],
    )
    vf_col = st.sidebar.selectbox(
        "[요약] VF 재고 열:", df_summary.columns, index=vf_idx
    )

    po_qty_idx = find_best_column(
        df_summary.columns,
        ["1회발주량", "1회발주", "1회", "PO수량", "발주단위", "파렛트", "배치"],
    )
    # 해당되는 컬럼이 없으면 (없음) 선택
    po_has_match = any(
        k in str(df_summary.columns[po_qty_idx])
        for k in ["1회", "발주", "PO", "단위", "파렛트"]
    )
    po_options = ["(없음 - 출고량 기반 계산)"] + list(df_summary.columns)
    po_qty_col = st.sidebar.selectbox(
        "[요약] 1회 발주량 열:",
        po_options,
        index=(po_qty_idx + 1) if po_has_match else 0,
    )

    inc_idx = find_best_column(
        df_summary.columns,
        ["입고예정", "매입예정", "입고대기", "입고예정량", "발주예정"],
    )
    inc_has_match = any(
        k in str(df_summary.columns[inc_idx])
        for k in ["입고", "매입", "대기", "예정"]
    ) and df_summary.columns[inc_idx] != (
        po_qty_col if po_has_match else ""
    )
    inc_options = ["(없음 - 0으로 계산)"] + list(df_summary.columns)
    incoming_col = st.sidebar.selectbox(
        "[요약] 입고 예정량 열:",
        inc_options,
        index=(inc_idx + 1) if inc_has_match else 0,
    )

    # 2) 출고 시트 열 자동 매칭
    sku_out_idx = find_best_column(
        df_daily_out.columns,
        ["품명", "상품명", "SKU", "품목", "상품", "품목명"],
    )
    sku_out_col = st.sidebar.selectbox(
        "[출고] 품명/SKU 열:", df_daily_out.columns, index=sku_out_idx
    )

    type_out_idx = find_best_column(
        df_daily_out.columns, ["유형", "카테고리", "분류", "구분", "종류"]
    )
    type_out_col = st.sidebar.selectbox(
        "[출고] 유형(카테고리) 열:", df_daily_out.columns, index=type_out_idx
    )

    # 출고 일자 컬럼 자동 분리
    non_date_cols = [sku_out_col, type_out_col, "평균", "합계", "비고"]
    date_cols = [
        c
        for c in df_daily_out.columns
        if c not in non_date_cols and not str(c).startswith("Unnamed")
    ]

    # 데이터 정제
    df_summary = df_summary[
        df_summary[sku_col].notna() & (df_summary[sku_col] != "nan")
    ]
    df_daily_out = df_daily_out[
        df_daily_out[sku_out_col].notna() & (df_daily_out[sku_out_col] != "nan")
    ]

    df_summary[main_col] = (
        pd.to_numeric(df_summary[main_col], errors="coerce").fillna(0).astype(int)
    )
    df_summary[vf_col] = (
        pd.to_numeric(df_summary[vf_col], errors="coerce").fillna(0).astype(int)
    )

    if po_qty_col != "(없음 - 출고량 기반 계산)":
      df_summary["기본_1회발주량"] = (
          pd.to_numeric(df_summary[po_qty_col], errors="coerce")
          .fillna(0)
          .astype(int)
      )
    else:
      df_summary["기본_1회발주량"] = 0

    if incoming_col != "(없음 - 0으로 계산)":
      df_summary["입고예정량"] = (
          pd.to_numeric(df_summary[incoming_col], errors="coerce")
          .fillna(0)
          .astype(int)
      )
    else:
      df_summary["입고예정량"] = 0

    for c in date_cols:
      df_daily_out[c] = pd.to_numeric(df_daily_out[c], errors="coerce")

    # 최근 일평균 출고량 계산
    df_daily_out["계산_일평균출고"] = (
        df_daily_out[date_cols].mean(axis=1, skipna=True).fillna(0)
    )

    # 데이터 병합
    merged = pd.merge(
        df_summary,
        df_daily_out[[sku_out_col, type_out_col, "계산_일평균출고"]],
        left_on=sku_col,
        right_on=sku_out_col,
        how="left",
    ).fillna({"계산_일평균출고": 0, type_out_col: "기타"})

    # 재고 및 소진일 계산
    merged["현재가용재고"] = merged[main_col] + merged[vf_col]
    merged["입고반영_총재고"] = merged["현재가용재고"] + merged["입고예정량"]

    merged["현재재고_소진일"] = np.where(
        merged["계산_일평균출고"] > 0,
        merged["현재가용재고"] / merged["계산_일평균출고"],
        999,
    )
    merged["입고반영_최종소진일"] = np.where(
        merged["계산_일평균출고"] > 0,
        merged["입고반영_총재고"] / merged["계산_일평균출고"],
        999,
    )
    merged["발주여유일(D-Day)"] = (
        merged["입고반영_최종소진일"] - lead_time - safety_days
    )

    today = datetime.today().date()

    def evaluate_status(d_day, cur_days, in_qty):
      if cur_days <= lead_time and in_qty == 0:
        return "🚨 [초긴급] 품절위험 (즉시발주)"
      elif cur_days <= lead_time and in_qty > 0:
        return f"🚚 입고대기중 (차기발주 D-{max(int(d_day), 0)})"
      elif d_day <= 0:
        return "🔥 [긴급] 오늘 차기발주 필요"
      elif d_day <= 7:
        return f"⚠️ [준비] 차기발주 D-{int(d_day)}"
      else:
        return f"✅ [안정] 차기발주 D-{int(d_day)}"

    merged["발주상태"] = merged.apply(
        lambda r: evaluate_status(
            r["발주여유일(D-Day)"],
            r["현재재고_소진일"],
            r["입고예정량"],
        ),
        axis=1,
    )

    # 2. 상단 카테고리(유형) 필터
    st.divider()
    all_types = ["전체 보기"] + list(merged[type_out_col].dropna().unique())
    selected_type = st.selectbox(
        "📂 조회할 품목 유형(카테고리)을 선택하세요:", all_types
    )

    filtered_df = (
        merged
        if selected_type == "전체 보기"
        else merged[merged[type_out_col] == selected_type]
    )

    # 3. 요약 KPI 카드
    urgent_cnt = len(
        filtered_df[filtered_df["발주상태"].str.contains("초긴급|오늘")]
    )
    inbound_cnt = len(
        filtered_df[filtered_df["발주상태"].str.contains("입고대기중")]
    )
    warning_cnt = len(filtered_df[filtered_df["발주상태"].str.contains("준비")])

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("선택 유형", f"{selected_type}")
    k2.metric("🚨 즉시 발주 필요", f"{urgent_cnt} 개 SKU")
    k3.metric("🚚 입고 대기 진행 중", f"{inbound_cnt} 개 SKU")
    k4.metric("⚠️ 7일 내 차기발주", f"{warning_cnt} 개 SKU")

    # 4. 유형별 전체 품목 일괄 비교 테이블
    st.subheader(f"📋 [{selected_type}] 전체 품목 발주 현황 일괄 비교표")

    view_cols = [
        sku_col,
        type_out_col,
        main_col,
        vf_col,
        "기본_1회발주량",
        "입고예정량",
        "입고반영_총재고",
        "계산_일평균출고",
        "현재재고_소진일",
        "입고반영_최종소진일",
        "발주상태",
    ]
    view_df = filtered_df[view_cols].copy()
    view_df = view_df.rename(
        columns={
            sku_col: "품명",
            type_out_col: "유형",
            main_col: "본창고재고",
            vf_col: "VF재고",
            "기본_1회발주량": "1회발주량(기준)",
            "계산_일평균출고": "일평균출고량",
            "현재재고_소진일": "현재재고 소진일",
            "입고반영_최종소진일": "입고반영 최종소진일",
        }
    )

    view_df["현재재고 소진일"] = view_df["현재재고 소진일"].apply(
        lambda x: f"{x:.1f}일" if x < 900 else "내역없음"
    )
    view_df["입고반영 최종소진일"] = view_df["입고반영 최종소진일"].apply(
        lambda x: f"{x:.1f}일" if x < 900 else "내역없음"
    )
    view_df["일평균출고량"] = view_df["일평균출고량"].apply(
        lambda x: f"{x:.1f}개/일"
    )

    st.dataframe(view_df.sort_values(by="발주상태"), use_container_width=True)

    # 5. 하단: 개별 SKU 선택 후 정밀 쿠팡 PO & 계단식 차감 시뮬레이션
    st.divider()
    st.subheader("🔍 개별 SKU 쿠팡 PO 연동 & 본창고 계단식 차감 시뮬레이터")

    target_sku = st.selectbox(
        "정밀 시뮬레이션할 품목을 선택하세요:",
        filtered_df[sku_col].unique(),
    )
    item = filtered_df[filtered_df[sku_col] == target_sku].iloc[0]

    base_po = int(item["기본_1회발주량"])
    if base_po == 0:
      base_po = (
          int(item["계산_일평균출고"] * 3)
          if item["계산_일평균출고"] > 0
          else 320
      )

    p1, p2, p3, p4 = st.columns(4)
    with p1:
      vf_trigger = st.number_input(
          "쿠팡 발주 트리거 (VF 잔여 기준)",
          min_value=0,
          value=200,
          step=10,
          help="VF 재고가 이 수량 미만이면 쿠팡 PO 발생",
      )
    with p2:
      po_multiplier = st.selectbox(
          "쿠팡 PO 발주 배수 (성수기/급증 대응)",
          [1, 2, 3],
          format_func=lambda x: f"{x}배수 ({base_po * x:,}개)",
          help="평상시는 1배수, 행사/성수기에는 2배수로 발주가 터질 때를 시뮬레이션합니다.",
      )
      po_batch_qty = base_po * po_multiplier

    with p3:
      sim_adu = st.number_input(
          "일평균 고객 출고량 (개/일)",
          min_value=0.0,
          value=float(item["계산_일평균출고"]),
          step=1.0,
      )
    with p4:
      incoming_val = st.number_input(
          "입고 예정 수량 (개)",
          min_value=0,
          value=int(item["입고예정량"]),
          step=100,
      )
      incoming_eta_days = st.number_input(
          "입고 예정일 (며칠 뒤 입고?)",
          min_value=0,
          max_value=90,
          value=7,
      )

    # 90일 정밀 시뮬레이션 연산
    sim_days = 90
    sim_dates = [today + timedelta(days=i) for i in range(sim_days)]

    curr_main = item[main_col]
    curr_vf = item[vf_col]

    hist_main = []
    hist_vf = []
    po_log = []
    main_stockout_day = None
    first_po_date = None

    for day_idx in range(sim_days):
      if incoming_val > 0 and day_idx == incoming_eta_days:
        curr_main += incoming_val
        po_log.append((
            sim_dates[day_idx],
            f"🚚 [본창고 입고완료] +{incoming_val:,}개 충전",
            curr_main,
            curr_vf,
        ))

      if sim_adu > 0:
        curr_vf -= sim_adu

      if sim_adu > 0 and curr_vf < vf_trigger:
        if first_po_date is None:
          first_po_date = sim_dates[day_idx]

        if curr_main >= po_batch_qty:
          curr_main -= po_batch_qty
          curr_vf += po_batch_qty
          po_log.append((
              sim_dates[day_idx],
              f"📦 쿠팡 PO 발주 -> 본창고 {po_batch_qty:,}개 차감 (정상 이동)",
              curr_main,
              curr_vf,
          ))
        else:
          if main_stockout_day is None:
            main_stockout_day = day_idx
          po_log.append((
              sim_dates[day_idx],
              (
                  f"🚨 [본창고 결품!] PO {po_batch_qty:,}개 중 잔여 {curr_main:,}개만"
                  " 이동"
              ),
              0,
              curr_vf + curr_main,
          ))
          curr_vf += curr_main
          curr_main = 0

      hist_main.append(curr_main)
      hist_vf.append(max(curr_vf, 0))

    # 본창고 결품일 및 발주 데드라인 계산
    if sim_adu == 0:
      stockout_str = "출고량 없음 (안전)"
      stockout_delta = None
      deadline_str = "발주 불필요"
      d_day_sim = 999
    elif main_stockout_day is None:
      stockout_str = "90일 이상 안전"
      stockout_delta = None
      deadline_str = "90일 이후"
      d_day_sim = 999
    else:
      stockout_date = sim_dates[main_stockout_day]
      stockout_str = f"{stockout_date.strftime('%Y-%m-%d')}"
      stockout_delta = f"D+{main_stockout_day}일 고갈"
      order_deadline = stockout_date - timedelta(days=(lead_time + safety_days))
      deadline_str = order_deadline.strftime("%Y-%m-%d")
      d_day_sim = (order_deadline - today).days

    # 실시간 연산 결과 카드
    st.markdown("---")
    st.markdown("#### 🎯 실시간 시뮬레이션 연산 결과")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric(
        "시작 가용 재고",
        f"총 {item[main_col] + item[vf_col]:,}개",
        f"본창고 {item[main_col]:,} / VF {item[vf_col]:,}",
    )
    m2.metric(
        "다음 쿠팡 PO 예상일",
        f"{first_po_date.strftime('%Y-%m-%d')}"
        if first_po_date
        else "발주조건 미달",
        f"적용 PO 단위: {po_batch_qty:,}개 ({po_multiplier}배수)",
    )
    m3.metric(
        "본창고 결품(재고소진) 예정일",
        f"{stockout_str}",
        stockout_delta,
        delta_color="inverse",
    )
    m4.metric(
        "구매팀 발주 요청 데드라인",
        f"{deadline_str}",
        f"D-Day {d_day_sim}일" if d_day_sim < 900 else None,
        delta_color="inverse",
    )

    if sim_adu > 0:
      if d_day_sim <= 0:
        st.error(
            f"🚨 **[초긴급 구매발주]** 리드타임({lead_time}일) 감안 시, **오늘 즉시 구매팀에 발주 요청**해야 {stockout_str} 본창고 결품을 막을 수 있습니다!"
        )
      elif d_day_sim <= 7:
        st.warning(
            f"⚠️ **[발주 준비 구간]** 구매팀 발주 데드라인({deadline_str})까지 **{d_day_sim}일** 남았습니다. 사전 검토를 진행하세요."
        )
      else:
        st.success(
            f"✅ **[재고 안정]** 입고 예정량({incoming_val:,}개) 반영 시 다음 발주 데드라인({deadline_str})까지 **{d_day_sim}일의 여유**가 있습니다."
        )
    else:
      st.info(
          "💡 일평균 출고량이 0개로 설정되어 있어 재고가 소진되지 않습니다. 출고량을 입력해 보세요."
      )

    # 계단식 그래프 출력
    st.markdown(
        f"#### 📈 [{target_sku}] 향후 60일간 본창고 vs VF 재고 변화 예측"
    )
    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=sim_dates[:60],
            y=hist_main[:60],
            mode="lines+markers",
            name="본창고 재고 (쿠팡 PO 시 계단식 차감)",
            line=dict(color="#1f77b4", width=3),
        )
    )
    fig.add_trace(
        go.Scatter(
            x=sim_dates[:60],
            y=hist_vf[:60],
            mode="lines",
            name="VF 잔여 재고 (일일 출고로 감소)",
            line=dict(color="#ff7f0e", width=2, dash="dash"),
        )
    )
    fig.add_hline(
        y=vf_trigger,
        line_dash="dot",
        line_color="red",
        annotation_text=f"쿠팡 발주 트리거 ({vf_trigger}개)",
    )

    fig.update_layout(
        xaxis_title="날짜",
        yaxis_title="재고 수량 (개)",
        hovermode="x unified",
        margin=dict(l=20, r=20, t=30, b=20),
    )
    st.plotly_chart(fig, use_container_width=True)

    with st.expander("📅 일자별 쿠팡 PO 발생 및 본창고 이동/입고 타임라인"):
      if po_log:
        log_df = pd.DataFrame(
            po_log,
            columns=[
                "예상일자",
                "이벤트 내용",
                "본창고 잔여",
                "VF 잔여",
            ],
        )
        st.dataframe(log_df, use_container_width=True)

  except Exception as e:
    st.error(f"오류가 발생했습니다: {e}")
else:
  st.info("👈 왼쪽 사이드바에서 엑셀 파일을 업로드해 주세요.")