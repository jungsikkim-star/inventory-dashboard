from datetime import datetime, timedelta
import re
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    layout="wide", page_title="쿠팡 VF & 평택(본창고) 통합 발주 대시보드"
)
st.title("📦 쿠팡 VF & 평택물류 통합 마스터 발주 시스템")
st.caption(
    "일일 출고량 + 쿠팡 납품 MOQ + 평택/VF 실재고 + 우측 입고 스케줄을 통합하여 안전재고 기반 발주 데드라인을 산출합니다."
)

# 1. 사이드바 설정
st.sidebar.header("📁 엑셀 파일 업로드")
uploaded_file = st.sidebar.file_uploader(
    "통합 마스터 엑셀 업로드 (.xlsx)", type=["xlsx"]
)

st.sidebar.divider()
st.sidebar.header("⚙️ 기본 설정")
base_date_input = st.sidebar.date_input(
    "재고 기준일자", datetime(2026, 9, 15).date()
)
default_lead_time = st.sidebar.number_input(
    "기본 구매 리드타임 (일)", min_value=1, max_value=120, value=35
)
vf_trigger_default = st.sidebar.number_input(
    "기본 쿠팡 PO 트리거 (VF 잔여)", min_value=0, value=200, step=10
)


# 헤더 텍스트 클리닝 헬퍼
def clean_txt(val):
  if pd.isna(val):
    return ""
  return str(val).strip().replace("\n", " ").replace(" ", "")


if uploaded_file:
  try:
    xl = pd.ExcelFile(uploaded_file)
    sheet_name = st.sidebar.selectbox("조회할 시트 선택:", xl.sheet_names, index=0)

    # 헤더 3행을 분석하기 위해 상단 데이터 로드 (header=None)
    raw_df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)

    # 1행~3행 헤더 스캔하여 row 1과 row 2 합성 (0-indexed 기준 row 1 = 엑셀 2행, row 2 = 엑셀 3행)
    header_row_main = 2  # 엑셀의 3행 (0901, 품명 등이 있는 행)
    for r in range(min(5, len(raw_df))):
      row_str = " ".join([clean_txt(x) for x in raw_df.iloc[r].values])
      if "0901" in row_str or "품명" in row_str:
        header_row_main = r
        break

    row_upper = (
        raw_df.iloc[header_row_main - 1] if header_row_main > 0 else None
    )
    row_lower = raw_df.iloc[header_row_main]

    combined_cols = []
    for idx in range(raw_df.shape[1]):
      u_val = clean_txt(row_upper[idx]) if row_upper is not None else ""
      l_val = clean_txt(row_lower[idx])
      comb = f"{u_val}_{l_val}".strip("_")
      combined_cols.append(comb if comb else f"COL_{idx}")

    df_data = raw_df.iloc[header_row_main + 1 :].copy()
    df_data.columns = combined_cols

    # 핵심 컬럼 자동 식별
    sku_col = None
    category_col = None
    moq_col = None
    vf_col = None
    main_col = None
    safety_col = None

    date_cols = []
    inbound_cols = []  # (컬럼명, 입고월, 입고일)

    for c in df_data.columns:
      c_clean = c.replace(" ", "")

      # 품명 컬럼
      if "품명" in c_clean and not sku_col:
        sku_col = c
      # 구분/유형
      elif any(k in c_clean for k in ["구분", "유형", "카테고리"]) and not category_col:
        category_col = c
      # 납품 MOQ
      elif "MOQ" in c_clean.upper() or "1회" in c_clean:
        moq_col = c
      # VF 재고
      elif "VF" in c_clean.upper() and ("재고" in c_clean or "수량" in c_clean):
        vf_col = c
      # 본창고/평택 재고
      elif (
          any(k in c_clean for k in ["평택", "본창고", "본물류"])
          and "재고" in c_clean
      ):
        main_col = c
      # 안전재고
      elif "안전재고" in c_clean:
        safety_col = c

      # 0901, 0902 등 일자별 출고 컬럼 (4자리 숫자 패턴)
      if re.search(r"\b09\d{2}\b", c_clean) or re.search(r"_\d{4}$", c_clean):
        date_cols.append(c)

      # 우측 입고 컬럼 (헤더에 '입고일:' 또는 '월'이 들어있는 패턴)
      if "입고일" in c or "입고" in c:
        match = re.search(r"(\d{1,2})월\s*(\d{1,2})일", c)
        if match:
          inbound_cols.append((c, int(match.group(1)), int(match.group(2))))

    # 품명 컬럼이 없는 경우 대비 보정
    if not sku_col:
      sku_col = df_data.columns[1]

    # 결측 행 제거 (품명이 없거나 합계 행 제거)
    df_data = df_data[df_data[sku_col].notna()]
    df_data = df_data[
        ~df_data[sku_col].astype(str).str.contains("합계|평균|비고", na=False)
    ]

    # 숫자 데이터 정제 함수
    def clean_num(val):
      if pd.isna(val) or str(val).strip() in ["-", "", "#DIV/0!", "전량"]:
        return 0
      try:
        return float(str(val).replace(",", "").strip())
      except:
        return 0

    # 숫자형 변환
    if moq_col:
      df_data[moq_col] = df_data[moq_col].apply(clean_num)
    if vf_col:
      df_data[vf_col] = df_data[vf_col].apply(clean_num)
    if main_col:
      df_data[main_col] = df_data[main_col].apply(clean_num)
    if safety_col:
      df_data[safety_col] = df_data[safety_col].apply(clean_num)

    # 일일 출고량 숫자 변환 및 일평균 출고량(ADU) 계산
    for dc in date_cols:
      df_data[dc] = df_data[dc].apply(clean_num)

    if date_cols:
      df_data["일평균출고"] = df_data[date_cols].mean(axis=1)
    else:
      df_data["일평균출고"] = 0.0

    # 우측 입고 컬럼 숫자 변환
    for ic, m, d in inbound_cols:
      df_data[ic] = df_data[ic].apply(clean_num)

    # 2. 상단 카테고리(유형) 필터
    st.divider()
    if category_col:
      all_types = ["전체 보기"] + list(
          df_data[category_col].dropna().unique()
      )
      selected_type = st.selectbox(
          "📂 조회할 구분(유형) 선택:", all_types
      )
      if selected_type != "전체 보기":
        filtered_df = df_data[df_data[category_col] == selected_type]
      else:
        filtered_df = df_data
    else:
      filtered_df = df_data

    # 전체 요약 표 계산
    records = []
    for idx, row in filtered_df.iterrows():
      sku_name = str(row[sku_col]).strip()
      moq = int(row[moq_col]) if moq_col else 100
      vf_stock = int(row[vf_col]) if vf_col else 0
      main_stock = int(row[main_col]) if main_col else 0
      safety_stock = int(row[safety_col]) if safety_col else 0
      adu = float(row["일평균출고"])

      # 총 가용 재고
      tot_stock = vf_stock + main_stock

      # 안전재고 적용 여부 (숫자가 0보다 크면 안전재고 기준, 0이면 원래대로 품절 0개 기준)
      has_safety = safety_stock > 0
      target_limit = safety_stock if has_safety else 0

      # 단순 계산 일수
      days_to_target = (
          max((tot_stock - target_limit) / adu, 0.0) if adu > 0 else 999
      )
      d_day = days_to_target - default_lead_time

      if tot_stock <= target_limit:
        status = "🚨 [초긴급] 마지노선 붕괴 (즉시발주)"
      elif d_day <= 0:
        status = "🔥 [긴급] 오늘 발주 필요"
      elif d_day <= 7:
        status = f"⚠️ [주의] D-{int(d_day)}일 내 발주"
      else:
        status = f"✅ [안정] D-{int(d_day)}일 여유"

      records.append({
          "품명": sku_name,
          "평택(본창고)": main_stock,
          "VF재고": vf_stock,
          "납품MOQ": moq,
          "안전재고": f"{safety_stock:,}개" if has_safety else "미적용",
          "일출고량": f"{adu:.1f}개/일",
          "목표도달소진일": (
              f"{days_to_target:.1f}일" if days_to_target < 900 else "-"
          ),
          "발주상태": status,
          "raw_sku": sku_name,
          "d_day_num": d_day,
      })

    summary_res_df = pd.DataFrame(records)

    # KPI 지표
    u_cnt = len(summary_res_df[summary_res_df["발주상태"].str.contains("초긴급|오늘")])
    w_cnt = len(summary_res_df[summary_res_df["발주상태"].str.contains("주의")])

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("기준 일자", base_date_input.strftime("%Y-%m-%d"))
    k2.metric("🚨 즉시 발주 필요", f"{u_cnt} 개 SKU")
    k3.metric("⚠️ 이번 주 발주 검토", f"{w_cnt} 개 SKU")
    k4.metric("📦 총 관리 SKU", f"{len(summary_res_df)} 개")

    st.subheader("📋 전체 품목 발주 현황 마스터 테이블")
    st.dataframe(
        summary_res_df.drop(columns=["raw_sku", "d_day_num"]).sort_values(
            by="발주상태"
        ),
        use_container_width=True,
    )

    # 3. 개별 품목 정밀 시뮬레이션 (우측 입고일정 완벽 반영)
    st.divider()
    st.subheader(
        "🔍 개별 SKU 정밀 시뮬레이터 (쿠팡 PO 연동 + 우측 입고 스케줄 자동 반영)"
    )

    selected_target = st.selectbox(
        "정밀 시뮬레이션할 품목 선택:", summary_res_df["raw_sku"].unique()
    )
    target_row = filtered_df[
        filtered_df[sku_col].astype(str).str.strip() == selected_target
    ].iloc[0]

    base_moq = int(target_row[moq_col]) if moq_col else 100
    if base_moq <= 0:
      base_moq = 100

    sim_c1, sim_c2, sim_c3, sim_c4 = st.columns(4)
    with sim_c1:
      sim_vf_trig = st.number_input(
          "쿠팡 발주 트리거 (VF 잔여)",
          min_value=0,
          value=int(vf_trigger_default),
          step=10,
      )
    with sim_c2:
      sim_multiplier = st.selectbox(
          "쿠팡 PO 발주 배수",
          [1, 2, 3],
          format_func=lambda x: f"{x}배수 ({base_moq * x:,}개)",
      )
      cur_po_unit = base_moq * sim_multiplier
    with sim_c3:
      sim_adu_input = st.number_input(
          "일평균 고객 출고량 (개/일)",
          min_value=0.0,
          value=float(target_row["일평균출고"]),
          step=1.0,
      )
    with sim_c4:
      sim_safety_input = st.number_input(
          "안전재고 수량 (0이면 원래대로)",
          min_value=0,
          value=int(target_row[safety_col]) if safety_col else 0,
          step=100,
      )

    # 해당 품목의 우측 시트 입고 스케줄 추출
    sku_inbound_schedule = []  # (입고날짜 datetime.date, 수량)
    for ic, m, d in inbound_cols:
      val = target_row[ic]
      if val > 0:
        # 연도 추정 (9월 이전이면 2027년, 9월 이후면 2026년)
        year = (
            base_date_input.year
            if m >= base_date_input.month
            else base_date_input.year + 1
        )
        in_date = datetime(year, m, d).date()
        if in_date >= base_date_input:
          sku_inbound_schedule.append((in_date, int(val)))

    sku_inbound_schedule.sort(key=lambda x: x[0])

    # 90일 시뮬레이션
    sim_days = 90
    sim_dates = [base_date_input + timedelta(days=i) for i in range(sim_days)]

    curr_main = int(target_row[main_col]) if main_col else 0
    curr_vf = int(target_row[vf_col]) if vf_col else 0

    hist_main = []
    hist_vf = []
    po_logs = []

    target_breach_day = None
    first_po_date = None

    for day_idx in range(sim_days):
      today_date = sim_dates[day_idx]

      # 1) 우측 스케줄상 오늘 입고되는 물량이 있다면 평택(본창고)에 자동 충전
      for in_date, in_qty in sku_inbound_schedule:
        if in_date == today_date:
          curr_main += in_qty
          po_logs.append((
              today_date,
              f"🚚 [입고 스케줄 반영] +{in_qty:,}개 평택본창고 입고 완료",
              curr_main,
              curr_vf,
          ))

      # 2) 일일 출고
      if sim_adu_input > 0:
        curr_vf -= sim_adu_input

      # 3) 쿠팡 PO 발생 체크
      if sim_adu_input > 0 and curr_vf < sim_vf_trig:
        if first_po_date is None:
          first_po_date = today_date

        if curr_main >= cur_po_unit:
          curr_main -= cur_po_unit
          curr_vf += cur_po_unit

          # 안전재고가 있는 경우 안전재고 하향 돌파 감지
          if (
              sim_safety_input > 0
              and curr_main <= sim_safety_input
              and target_breach_day is None
          ):
            target_breach_day = day_idx

          po_logs.append((
              today_date,
              (
                  f"📦 쿠팡 PO {cur_po_unit:,}개 발생 ➔ 평택창고 차감 (잔여:"
                  f" {curr_main:,}개)"
              ),
              curr_main,
              curr_vf,
          ))
        else:
          # 평택창고 결품!
          if target_breach_day is None:
            target_breach_day = day_idx

          po_logs.append((
              today_date,
              (
                  f"🚨 [평택창고 결품/부족!] PO {cur_po_unit:,}개 중 잔여 {curr_main:,}개만"
                  " VF로 이동"
              ),
              0,
              curr_vf + curr_main,
          ))
          curr_vf += curr_main
          curr_main = 0

      hist_main.append(curr_main)
      hist_vf.append(max(curr_vf, 0))

    # 결과 산출
    is_safety_mode = sim_safety_input > 0
    mode_name = "안전재고선 붕괴일" if is_safety_mode else "평택창고 품절(0개)일"

    if sim_adu_input == 0:
      breach_str = "출고량 없음"
      deadline_str = "발주 불필요"
      d_day_num = 999
    elif target_breach_day is None:
      breach_str = "90일 이상 안전"
      deadline_str = "여유"
      d_day_num = 999
    else:
      b_date = sim_dates[target_breach_day]
      breach_str = f"{b_date.strftime('%Y-%m-%d')} (D+{target_breach_day}일)"
      d_line = b_date - timedelta(days=default_lead_time)
      deadline_str = d_line.strftime("%Y-%m-%d")
      d_day_num = (d_line - base_date_input).days

    # 실시간 카드
    st.markdown("---")
    res1, res2, res3, res4 = st.columns(4)
    res1.metric(
        "기준 재고 (09/15)",
        f"평택 {int(target_row[main_col]):,}개 / VF {int(target_row[vf_col]):,}개",
        f"안전재고: {sim_safety_input:,}개" if is_safety_mode else "미적용",
    )
    res2.metric(
        "다음 쿠팡 PO 예상일",
        f"{first_po_date.strftime('%m/%d')}" if first_po_date else "조건 미달",
        f"1회 납품MOQ: {cur_po_unit:,}개",
    )
    res3.metric(
        f"{mode_name}",
        f"{breach_str}",
        f"예정 입고 {len(sku_inbound_schedule)}건 스케줄 반영됨",
    )
    res4.metric(
        "구매팀 발주 요청 데드라인",
        f"{deadline_str}",
        f"D-Day {d_day_num}일" if d_day_num < 900 else None,
        delta_color="inverse",
    )

    # 안내 배너
    if sim_adu_input > 0:
      if d_day_num <= 0:
        st.error(
            f"🚨 **[초긴급 구매발주]** 리드타임({default_lead_time}일) 고려 시, **오늘 즉시 발주 요청**해야 {mode_name}({breach_str})을 막을 수 있습니다!"
        )
      elif d_day_num <= 7:
        st.warning(
            f"⚠️ **[발주 준비]** 구매 데드라인({deadline_str})까지 **{d_day_num}일** 남았습니다."
        )
      else:
        st.success(
            f"✅ **[안정]** 예정된 입고 스케줄 덕분에 다음 발주 데드라인({deadline_str})까지 **{d_day_num}일의 여유**가 있습니다."
        )

    # 계단식 그래프
    st.markdown(
        f"#### 📈 [{selected_target}] 향후 60일 재고 시뮬레이션 (입고 자동 충전 반영)"
    )
    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=sim_dates[:60],
            y=hist_main[:60],
            mode="lines+markers",
            name="평택본창고 (쿠팡 PO 시 계단식 차감 & 입고일 충전)",
            line=dict(color="#1f77b4", width=3),
        )
    )
    fig.add_trace(
        go.Scatter(
            x=sim_dates[:60],
            y=hist_vf[:60],
            mode="lines",
            name="VF 잔여 재고 (일일 고객출고)",
            line=dict(color="#ff7f0e", width=2, dash="dash"),
        )
    )

    if is_safety_mode:
      fig.add_hline(
          y=sim_safety_input,
          line_dash="dashdot",
          line_color="#ffbb00",
          annotation_text=f"안전재고 마지노선 ({sim_safety_input:,}개)",
      )

    fig.add_hline(
        y=sim_vf_trig,
        line_dash="dot",
        line_color="red",
        annotation_text=f"쿠팡 발주 트리거 ({sim_vf_trig}개)",
    )

    fig.update_layout(
        xaxis_title="날짜",
        yaxis_title="수량 (개)",
        hovermode="x unified",
        margin=dict(l=20, r=20, t=30, b=20),
    )
    st.plotly_chart(fig, use_container_width=True)

    # 타임라인 로그
    with st.expander("📅 일자별 상세 쿠팡 PO 차감 및 본창고 입고 타임라인"):
      if po_logs:
        st.dataframe(
            pd.DataFrame(
                po_logs,
                columns=["일자", "내용", "평택 잔여", "VF 잔여"],
            ),
            use_container_width=True,
        )
      else:
        st.info("시뮬레이션 기간 동안 발생한 이벤트가 없습니다.")

  except Exception as e:
    st.error(f"엑셀 데이터를 읽는 중 오류가 발생했습니다: {e}")
else:
  st.info("👈 왼쪽 사이드바에서 수정하신 통합 마스터 엑셀 파일을 업로드해 주세요.")
