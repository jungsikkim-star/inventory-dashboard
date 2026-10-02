"""
쿠팡 VF & 평택(본창고) 통합 마스터 발주 대시보드  (v3)

업로드 엑셀 구조
  - '쿠팡PO' 탭(맨 앞): 날짜 / 품명 / VF_매입(BOX) / 수량(PLT) / 단가 / 매출합계 / 구분(센터)
      → 헤더에 '날짜'와 '단가'가 있는 시트는 자동으로 PO 이력 시트로 인식합니다.
  - 월별 시트('9월', '10월' ...)
      상단 헤더 3행: [구분/품명/일자] [요일] [합계, MMDD 날짜]
      일자별 출고 → 납품 MOQ → VF재고 → 평택재고 → 평택 안전재고(45일분) → 발주필요
      그 오른쪽: '발주일 : ○월○일 / 입고일 : ○월○일' 형태의 공장 입고 스케줄 컬럼들

핵심 가정 (사이드바에서 조정 가능)
  - 구분이 '밀크런'인 품목은 VF재고가 없고, 일자별 출고 = 평택에서 쿠팡으로 직접 나간 물량.
  - 그 외(벤플/글로브 등)는 VF재고가 고객 출고로 줄고, VF가 트리거 밑으로 내려가면
    쿠팡 PO가 발생해 평택 재고가 MOQ × 배수만큼 VF로 이동.
  - 공장 발주 데드라인 = (안전재고 도달일 또는 평택 결품일) - 리드타임.

쿠팡 PO 예측 (v3 추가)
  - PO 1회 수량 = 최근 N회 PO(일자별 합) 수량의 중앙값
  - 출고율   = VF형: 시트의 최근 7일 출고 평균 / 밀크런: 최근 28일 PO 실적 ÷ 28
  - PO 간격  = PO 1회 수량 ÷ 출고율
  - 차기 PO  = VF형은 '마지막 PO 이후 실제 출고량'이 PO 1회 수량을 채우는 날,
               밀크런은 마지막 PO(이미 잡힌 미래 PO 포함) + PO 간격
  - 예상 매출 = PO 수량 × 최근 단가 (재고 제약 반영 시 평택 재고+발주완료 입고분 안에서만 납품)
"""
import calendar
import difflib
import hashlib
import math
import re
import statistics
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

KST = timezone(timedelta(hours=9))
HORIZON = 180          # 시뮬레이션 기간(일). 리드타임(최대 120일) 이후까지 볼 수 있어야 함
CHART_DAYS = 90
DEFAULT_MOQ = 100      # MOQ 미기재 시 임시값
PERIOD_DAY = {"초": 10, "중": 20, "말": None}   # '11월초/12월 중순/9월말' → 보수적으로 순의 끝날, 말=말일
FORECAST_MONTHS = 6    # PO/매출 예측 범위: 기준월 포함 6개월
NO_MATCH = "(매칭 안 함)"


# ─────────────────────────────── 공통 유틸 ───────────────────────────────
def today_kst() -> date:
    return datetime.now(KST).date()


def clean_txt(v) -> str:
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return ""
    return re.sub(r"\s+", "", str(v))


def to_num(v):
    """숫자면 float, 빈칸/문자(전량, O, X 등)면 None."""
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "").strip())
    except ValueError:
        return None


def to_date(v):
    """datetime / date / 엑셀 날짜 일련번호(46212 등) / 'YYYY-MM-DD' 문자열 → date."""
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)):
        if pd.isna(v) or not 20000 < v < 80000:
            return None
        return date(1899, 12, 30) + timedelta(days=int(v))
    m = re.match(r"\s*(\d{4})[-./](\d{1,2})[-./](\d{1,2})", str(v))
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    return None


def md_of(text: str):
    """'0901' / '901' → (9, 1). 날짜처럼 안 보이면 None."""
    if not text.isdigit() or not 3 <= len(text) <= 4:
        return None
    s = text.zfill(4)
    mo, dy = int(s[:2]), int(s[2:])
    return (mo, dy) if 1 <= mo <= 12 and 1 <= dy <= 31 else None


def add_months(y: int, m: int, k: int):
    t = y * 12 + (m - 1) + k
    return t // 12, t % 12 + 1


def month_end(y: int, m: int) -> date:
    return date(y, m, calendar.monthrange(y, m)[1])


_MD = re.compile(r"(\d{1,2})월(\d{1,2})일")
_SL = re.compile(r"(\d{1,2})[/.](\d{1,2})")
_MP = re.compile(r"(\d{1,2})월(초순|초|중순|중|하순|말)")


def parse_kr_date(text: str, anchor: date, not_before: date = None):
    """'9월4일' '10월 10일' '11월초' '12월 중순' '9월말' → (date, 추정여부). 연도는 anchor 근처로 추정."""
    t = clean_txt(text)
    est, dy = False, None
    m = _MD.search(t) or _SL.search(t)
    if m:
        mo, dy = int(m.group(1)), int(m.group(2))
    else:
        m = _MP.search(t)
        if not m:
            return None, False
        mo, kw, est = int(m.group(1)), m.group(2), True
        key = "초" if kw.startswith("초") else "중" if kw.startswith("중") else "말"
        dy = PERIOD_DAY[key]
    if not 1 <= mo <= 12:
        return None, False
    cands = []
    for y in (anchor.year - 1, anchor.year, anchor.year + 1):
        last = calendar.monthrange(y, mo)[1]
        cands.append(date(y, mo, min(dy or last, last)))
    if not_before:
        later = [d for d in cands if d >= not_before]
        return (min(later) if later else max(cands)), est
    return min(cands, key=lambda d: abs((d - anchor).days)), est


# ─────────────────────────────── 엑셀 파싱 ───────────────────────────────
def match_columns(headers):
    idx = {}
    for i, h in enumerate(headers):
        H = h.upper()
        if "품명" in h:
            idx.setdefault("name", i)
        elif "구분" in h or "유형" in h:
            idx.setdefault("cat", i)
        elif "MOQ" in H or "납품" in h:
            idx.setdefault("moq", i)
        elif "안전재고" in h:
            idx.setdefault("safety", i)
        elif "VF" in H and ("재고" in h or "수량" in h):
            idx.setdefault("vf", i)
        elif ("평택" in h or "본창고" in h or "본물류" in h) and "재고" in h:
            idx.setdefault("main", i)
    return idx


def parse_sheet(raw: pd.DataFrame, sheet_name: str, fallback_year: int) -> dict:
    n_rows, n_cols = raw.shape
    sku_row, best = None, (0, None)
    for r in range(min(8, n_rows)):
        vals = [clean_txt(x) for x in raw.iloc[r].tolist()]
        if sku_row is None and "품명" in vals:
            sku_row = r
        cnt = sum(1 for v in vals if md_of(v))
        if cnt > best[0]:
            best = (cnt, r)
    date_row = best[1] if best[0] >= 5 else None
    data_start = max(1 if sku_row is None else sku_row, 2 if date_row is None else date_row) + 1

    # 시트 제목('26년 9월')에서 연/월
    title = " ".join(clean_txt(raw.iat[r, c]) for r in range(min(3, n_rows)) for c in range(min(4, n_cols)))
    tm = re.search(r"(\d{2,4})년(\d{1,2})월", title)
    if tm:
        year = int(tm.group(1)) + (2000 if int(tm.group(1)) < 100 else 0)
        month = int(tm.group(2))
    else:
        mm = re.search(r"(\d{1,2})월", sheet_name)
        year, month = fallback_year, int(mm.group(1)) if mm else 0

    headers = []
    for c in range(n_cols):
        pieces = []
        for r in range(data_start):
            t = clean_txt(raw.iat[r, c])
            if t and t not in pieces:
                pieces.append(t)
        headers.append("_".join(pieces))
    cols = match_columns(headers)
    if "name" not in cols:
        raise ValueError(f"시트 '{sheet_name}'에서 '품명' 컬럼을 찾지 못했습니다.")

    # 기준일(헤더의 =TODAY() 캐시값)
    snapshot = None
    for key in ("vf", "main"):
        if key in cols:
            for r in range(data_start):
                v = raw.iat[r, cols[key]]
                if isinstance(v, date):
                    snapshot = v.date() if isinstance(v, datetime) else v
                    break
        if snapshot:
            break

    # 일자 컬럼
    date_cols = {}
    if date_row is not None:
        for c in range(n_cols):
            md = md_of(clean_txt(raw.iat[date_row, c]))
            if md:
                try:
                    date_cols[date(year, md[0], md[1])] = c
                except ValueError:      # 예: 9월 31일
                    pass

    # 입고 스케줄 컬럼 ('발주일 : ○ / 입고일 : ○')
    anchor = date(year, month or 1, 15)
    inbound_cols, unparsed = [], []
    for c in range(n_cols):
        pieces = [raw.iat[r, c] for r in range(data_start)]
        joined = "|".join(clean_txt(p) for p in pieces)
        if "입고일" not in joined:
            continue
        a_txt = re.search(r"입고일[:：]?([^|]*)", joined).group(1)
        o_m = re.search(r"발주일[:：]?([^|]*)", joined)
        order, o_est = (parse_kr_date(o_m.group(1), anchor) if o_m else (None, False))
        arrival, a_est = parse_kr_date(a_txt, anchor, not_before=order)
        group = next((re.sub(r"\s+", " ", str(p)).strip() for p in pieces
                      if clean_txt(p) and "발주일" not in clean_txt(p) and "입고일" not in clean_txt(p)), "미분류")
        if arrival is None:
            unparsed.append(f"{group} / {a_txt}")
            continue
        inbound_cols.append(dict(col=c, group=group, order=order, arrival=arrival, est=bool(o_est or a_est), arr_est=a_est,
                                 raw_order=(o_m.group(1) if o_m else ""), raw_arrival=a_txt))

    # 품목 행
    rows, valid_idx = [], []
    for r in range(data_start, n_rows):
        nm = raw.iat[r, cols["name"]]
        key = clean_txt(nm)
        if not key or key.lower() == "nan" or "합계" in key or key in ("평균", "비고"):
            continue
        cat_raw = raw.iat[r, cols["cat"]] if "cat" in cols else None
        if "합계" in clean_txt(cat_raw):
            continue
        valid_idx.append(r)

    entered = set()
    for d, c in date_cols.items():
        if any(pd.notna(raw.iat[r, c]) for r in valid_idx):
            entered.add(d)          # 한 품목이라도 값이 있으면 '입력된 날'

    for r in valid_idx:
        moq_raw = raw.iat[r, cols["moq"]] if "moq" in cols else None
        vf = to_num(raw.iat[r, cols["vf"]]) if "vf" in cols else None
        cat_raw = raw.iat[r, cols["cat"]] if "cat" in cols else None
        inbound = []
        for ic in inbound_cols:
            q = to_num(raw.iat[r, ic["col"]])
            if q and q > 0:
                inbound.append((ic["arrival"], q, ic["arr_est"], ic["group"], ic["order"], ic["col"]))
        rows.append(dict(
            key=clean_txt(raw.iat[r, cols["name"]]),
            name=str(raw.iat[r, cols["name"]]).strip(),
            cat=str(cat_raw).strip() if clean_txt(cat_raw) else "기타",
            moq=to_num(moq_raw), moq_all=clean_txt(moq_raw) == "전량",
            vf=vf,
            main=(to_num(raw.iat[r, cols["main"]]) or 0.0) if "main" in cols else 0.0,
            safety=(to_num(raw.iat[r, cols["safety"]]) or 0.0) if "safety" in cols else 0.0,
            sales={d: (to_num(raw.iat[r, c]) or 0.0) for d, c in date_cols.items() if d in entered},
            inbound=inbound,
        ))

    missing = [k for k in ("moq", "vf", "main") if k not in cols]
    return dict(name=sheet_name, year=year, month=month, snapshot=snapshot, rows=rows,
                entered=sorted(entered), inbound_cols=inbound_cols, unparsed=unparsed, missing=missing)


def is_po_sheet(raw: pd.DataFrame) -> bool:
    """헤더(앞 8행)에 '단가'와 '날짜/일자'가 같이 있으면 쿠팡 PO 이력 시트로 본다."""
    for r in range(min(8, raw.shape[0])):
        vals = {clean_txt(x) for x in raw.iloc[r].tolist()}
        if "단가" in vals and ("날짜" in vals or "일자" in vals):
            return True
    return False


def parse_po_sheet(raw: pd.DataFrame, sheet_name: str) -> dict:
    n_rows = raw.shape[0]
    hr = next(r for r in range(min(8, n_rows)) if "단가" in {clean_txt(x) for x in raw.iloc[r].tolist()})
    hdr = [clean_txt(x) for x in raw.iloc[hr].tolist()]

    def find(pred):
        return next((i for i, h in enumerate(hdr) if h and pred(h)), None)

    c_date = find(lambda h: h in ("날짜", "일자", "PO일자", "발주일"))
    c_name = find(lambda h: "품명" in h or "상품명" in h)
    c_qty = find(lambda h: "매입" in h or "BOX" in h.upper() or "확정수량" in h)
    if c_qty is None:
        c_qty = find(lambda h: "수량" in h and "PLT" not in h.upper())
    c_price = find(lambda h: "단가" in h)
    c_ctr = find(lambda h: h in ("구분", "센터", "입고센터", "물류센터"))
    miss = [k for k, v in (("날짜", c_date), ("품명", c_name), ("수량", c_qty), ("단가", c_price)) if v is None]
    if miss:
        raise ValueError(f"PO 시트에서 {', '.join(miss)} 컬럼을 찾지 못했습니다.")

    recs, bad_date = [], 0
    for r in range(hr + 1, n_rows):
        nm = raw.iat[r, c_name]
        if not clean_txt(nm):
            continue
        d = to_date(raw.iat[r, c_date])
        if d is None:
            bad_date += 1
            continue
        ctr = raw.iat[r, c_ctr] if c_ctr is not None else None
        recs.append(dict(date=d, name=re.sub(r"\s+", " ", str(nm)).strip(),
                         qty=to_num(raw.iat[r, c_qty]) or 0.0, price=to_num(raw.iat[r, c_price]) or 0.0,
                         center=re.sub(r"\s+", " ", str(ctr)).strip() if clean_txt(ctr) else ""))
    return dict(name=sheet_name, rows=recs, bad_date=bad_date)


PARSER_VERSION = 4     # 파싱 결과 구조를 바꿀 때마다 올리면 Streamlit이 예전 캐시를 버리고 다시 파싱함


@st.cache_data(show_spinner="엑셀 분석 중...")
def load_workbook(file_bytes: bytes, parser_version: int = PARSER_VERSION) -> dict:
    """월별 재고 시트와 쿠팡 PO 이력 시트를 나눠 읽는다. 형식이 맞지 않는 시트는 건너뛰고 이유를 남긴다."""
    import io
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    sheets, po, skipped = [], [], []
    for s in xl.sheet_names:
        raw = xl.parse(s, header=None)
        try:
            if is_po_sheet(raw):
                po.append(parse_po_sheet(raw, s))
            else:
                sheets.append(parse_sheet(raw, s, today_kst().year))
        except Exception as e:          # 메모 탭 등 형식이 다른 시트 때문에 전체가 멈추지 않도록
            skipped.append(f"{s}: {e}")
    return dict(sheets=sheets, po=po, skipped=skipped)


def as_date(v):
    if v is None or (not isinstance(v, (date, datetime)) and pd.isna(v)):
        return None
    return v.date() if isinstance(v, datetime) else (v if isinstance(v, date) else None)


def apply_inbound_overrides(sheets: list, sheet_name: str, overrides: dict) -> list:
    """overrides = {컬럼번호: (발주일 또는 None, 입고일)}. 화면에서 직접 고친 날짜는 '확정'으로 취급한다."""
    if not overrides:
        return sheets
    out = []
    for sh in sheets:
        if sh["name"] != sheet_name:
            out.append(sh)
            continue
        cols = []
        for c in sh["inbound_cols"]:
            o = overrides.get(c["col"])
            cols.append(dict(c, order=o[0], arrival=o[1], est=False, arr_est=False) if o else c)
        rows = []
        for r in sh["rows"]:
            ib = []
            for (arr, q, est, g, od, col) in r["inbound"]:
                o = overrides.get(col)
                ib.append((o[1], q, False, g, o[0], col) if o else (arr, q, est, g, od, col))
            rows.append(dict(r, inbound=ib))
        out.append(dict(sh, inbound_cols=cols, rows=rows))
    return out


# ─────────────────────────────── 품목 구성 ───────────────────────────────
@dataclass
class Item:
    name: str
    cat: str
    is_milkrun: bool
    moq: float
    moq_all: bool
    moq_missing: bool
    vf: float
    main: float
    safety: float
    adu: float
    adu_days: int
    lt: int
    lt_src: str
    inbound: list          # [(도착일, 수량, 추정여부)]  ← 발주일이 기준일 이전인 '발주 완료' 건만 (재고 계산에 반영)
    inbound_detail: list   # 기준일 이후 도착하는 전체 입고 건 [{arrival, qty, est, group, order, ordered}]
    in_latest: bool
    sheet: str
    plan: dict             # {날짜: 수량} 기준일 이후 시트에 미리 입력된 출고(쿠팡 PO 예정)
    plan_until: object     # 예정 출고가 입력된 마지막 날짜 (없으면 None)
    sales: dict            # {날짜: 출고} 전 월 시트 합본 (PO 예측에 사용)


def calc_adu(sales: dict, entered: list, base: date, window: int):
    """기준일 이전에 '입력된 날' 중 최근 window일의 평균 (빈 칸은 0, 미입력일은 제외)."""
    days = [d for d in entered if d < base][-window:]
    if not days:
        return 0.0, 0
    return sum(sales.get(d, 0.0) for d in days) / len(days), len(days)


def estimate_lead_times(sheets: list) -> dict:
    pairs = {}
    for sh in sheets:
        for c in sh["inbound_cols"]:
            if c["order"]:
                pairs.setdefault(c["group"], {})[(c["order"], c["arrival"])] = c["est"]
    out = {}
    for g, d in pairs.items():
        exact = [(a - o).days for (o, a), est in d.items() if not est]
        use = exact or [(a - o).days for (o, a) in d]
        out[g] = dict(median=int(statistics.median(use)), n=len(use), n_exact=len(exact))
    return out


def build_items(sheets, base, win_std, win_mr, default_lt, use_hist_lt, mr_keyword, include_stale, use_plan=True):
    sheets = sorted(sheets, key=lambda s: (s["year"], s["month"]))
    latest = sheets[-1]["name"]
    entered = sorted({d for s in sheets for d in s["entered"]})
    lt_est = estimate_lead_times(sheets)
    acc = {}
    for sh in sheets:                                   # 나중 월 시트가 재고/MOQ/입고 스케줄을 덮어씀
        for r in sh["rows"]:
            a = acc.setdefault(r["key"], dict(sales={}, groups=set()))
            a["sales"].update(r["sales"])
            a["groups"] |= {g for (_, _, _, g, _, _) in r["inbound"]}
            a["row"], a["sheet"] = r, sh["name"]
    items = []
    for a in acc.values():
        r, in_latest = a["row"], a["sheet"] == latest
        if not in_latest and not include_stale:
            continue
        is_mr = mr_keyword in r["cat"]
        adu, n = calc_adu(a["sales"], entered, base, win_mr if is_mr else win_std)
        lt, src = int(default_lt), "기본값"
        if use_hist_lt:
            c = [(lt_est[g]["median"], g, lt_est[g]["n"]) for g in a["groups"] if g in lt_est]
            if c:
                v, g, cnt = max(c)
                lt, src = v, f"실적 {g} (n={cnt})"
        moq_missing = not r["moq_all"] and not (r["moq"] and r["moq"] > 0)
        plan = {d: q for d, q in a["sales"].items() if d >= base and q > 0} if use_plan else {}
        detail = []
        for (d_arr, q_in, est_in, grp, d_ord, _col) in r["inbound"]:
            if d_arr < base:
                continue                                    # 이미 입고된 건은 현재 재고에 포함돼 있음
            ordered = d_ord is None or d_ord <= base        # 발주일이 기준일 이전이거나 당일이면 발주 완료 (미기재는 완료로 간주)
            detail.append(dict(arrival=d_arr, qty=q_in, est=est_in, group=grp, order=d_ord, ordered=ordered))
        detail.sort(key=lambda x: (x["arrival"], x["order"] or date.min))
        items.append(Item(
            name=r["name"], cat=r["cat"], is_milkrun=is_mr,
            moq=float(r["moq"]) if r["moq"] and r["moq"] > 0 else float(DEFAULT_MOQ),
            moq_all=r["moq_all"], moq_missing=moq_missing,
            vf=0.0 if is_mr else float(r["vf"] or 0.0), main=float(r["main"]), safety=float(r["safety"]),
            adu=adu, adu_days=n, lt=lt, lt_src=src,
            inbound=[(x["arrival"], x["qty"], x["est"]) for x in detail if x["ordered"]],
            inbound_detail=detail,
            in_latest=in_latest, sheet=a["sheet"],
            plan=plan, plan_until=max(plan) if plan else None,
            sales=dict(a["sales"]),
        ))
    return items, latest, entered, lt_est


# ─────────────────────────────── 시뮬레이션 ───────────────────────────────
def simulate(item: Item, base: date, adu: float, trig_mode: str, trig_value: float,
             po_mult: int, transfer_days: int, horizon: int = HORIZON, log: bool = False) -> dict:
    """
    무발주(신규 공장 발주 없음) 기준으로 평택/VF 재고를 일 단위로 굴린다.
      밀크런형: 평택이 일평균 출고만큼 직접 감소
      VF형    : VF가 출고로 감소 → VF(+이동중) < 트리거 이면 쿠팡 PO(MOQ×배수)만큼 평택→VF 이동
    """
    main, vf = item.main, item.vf
    inbound = {}
    for d, q, _ in item.inbound:
        if d >= base:
            inbound[d] = inbound.get(d, 0.0) + q
    unit = None if item.moq_all else item.moq * po_mult
    trig = trig_value * adu if trig_mode == "days" else float(trig_value)
    pipe, events = [], []
    res = dict(first_po=None, stockout=None, short_days=[], n_po=0)
    main_s, vf_s, lost_s, lost = [], [], [], 0.0
    out_s, prev_out = [], False

    for i in range(horizon):
        day = base + timedelta(days=i)
        short_now = False
        q = inbound.get(day, 0.0)
        if q:
            main += q
            if log:
                events.append((day, f"🚚 [평택 입고] +{q:,.0f}개", main, vf))
        if pipe:
            arr = sum(x for t, x in pipe if t <= i)
            if arr:
                vf += arr
                pipe = [(t, x) for t, x in pipe if t > i]
                if log:
                    events.append((day, f"📥 [VF 도착] +{arr:,.0f}개 (쿠팡 PO 이동분)", main, vf))
        demand = adu
        if item.plan_until and day <= item.plan_until:      # 예정 출고가 입력된 기간은 그 값을 그대로 사용
            demand = item.plan.get(day, 0.0)
            if log and demand > 0 and item.is_milkrun:
                events.append((day, f"📤 [쿠팡 밀크런 출고 예정] -{demand:,.0f}개", max(main - demand, 0.0), vf))
        if demand > 0:
            if item.is_milkrun:
                if main >= demand:
                    main -= demand
                else:
                    lost += demand - main
                    main = 0.0
                    short_now = True
                    res["short_days"].append(i)
                    if res["stockout"] is None:
                        res["stockout"] = i
                        if log:
                            events.append((day, "🚨 [평택 결품] 밀크런 출고 물량 부족", main, vf))
            else:
                if vf >= demand:
                    vf -= demand
                else:
                    lost += demand - vf
                    vf = 0.0
                transit = sum(x for _, x in pipe)
                short_today = False
                for _ in range(20):
                    if vf + transit >= trig or main <= 1e-9:
                        break
                    qty = main if item.moq_all else min(unit, main)
                    if not item.moq_all and main < unit:
                        short_today = True
                    main -= qty
                    if transfer_days > 0:
                        pipe.append((i + transfer_days, qty))
                        transit += qty
                    else:
                        vf += qty
                    res["n_po"] += 1
                    res["first_po"] = res["first_po"] or day
                    if log:
                        tag = "" if item.moq_all or qty >= unit else " (MOQ 미달, 잔량 전부)"
                        events.append((day, f"📦 쿠팡 PO: 평택 -{qty:,.0f}개 → VF{tag}", main, vf))
                    if item.moq_all:
                        break
                if vf + transit < trig and main <= 1e-9:
                    short_today = True
                if short_today:
                    short_now = True
                    res["short_days"].append(i)
                    if res["stockout"] is None:
                        res["stockout"] = i
                        if log:
                            events.append((day, "🚨 [평택 결품] 쿠팡 PO를 채울 재고 부족", main, vf))
        # 밀크런은 평택이 0인 동안(출고 없는 날 포함) 계속 '결품 상태'로 본다
        prev_out = short_now or (item.is_milkrun and prev_out and main <= 1e-9)
        out_s.append(prev_out)
        main_s.append(main)
        vf_s.append(vf)
        lost_s.append(lost)
    res.update(main=main_s, vf=vf_s, lost=lost_s, out=out_s, events=events)
    return res


def _runs(flags):
    """True 가 이어지는 구간들 → [(시작, 끝)]"""
    runs, i, n = [], 0, len(flags)
    while i < n:
        if flags[i]:
            j = i
            while j + 1 < n and flags[j + 1]:
                j += 1
            runs.append((i, j))
            i = j + 1
        else:
            i += 1
    return runs


def evaluate(item: Item, sim: dict, base: date, adu: float, safety: float, lt: int, target_days: int) -> dict:
    """
    시뮬레이션 결과 → 데드라인 / 상태 / 권장 발주량.

    핵심 원칙: 이미 발주한(발주완료) 입고분을 반영한 뒤에도 '끝까지 회복되지 않는' 부족만 신규 발주 대상으로 본다.
      · 영구 부족  : 시뮬레이션 끝(HORIZON)까지 계속 안전재고 이하이거나 결품인 구간 → 신규 발주 데드라인 = 시작일 − 리드타임
      · 일시 부족  : 이후 발주완료 입고분으로 회복되는 구간 → 신규 발주 대상이 아니라 '입고 전 부족'으로 따로 표시
                     (단, 일시 결품이 리드타임 이후까지 이어지면 신규 발주로 단축 가능하므로 발주 대상에 포함)
    """
    ms, n = sim["main"], len(sim["main"])
    out = sim.get("out") or [False] * n
    has_demand = adu > 0 or bool(item.plan)
    pre = "입고 반영 후 " if item.inbound else ""
    md = lambda i: fmt_md(base, i)

    hard = _runs(out)
    soft = _runs([m <= safety for m in ms]) if safety > 0 else []
    pers_h = hard[-1] if hard and hard[-1][1] == n - 1 else None
    pers_s = soft[-1] if soft and soft[-1][1] == n - 1 else None
    trans_h = [r for r in hard if r != pers_h]
    trans_s = [r for r in soft if r != pers_s]

    cands = []                                   # (신규 발주 데드라인(일), 우선순위, 사유)
    unavoid = pers_h if (pers_h and pers_h[0] < lt) else None
    if pers_h:
        cands.append((max(pers_h[0], lt) - lt, 0, f"{pre}{md(pers_h[0])} 결품"))
    for a, b in trans_h:
        if b >= lt:
            cands.append((max(a, lt) - lt, 0, f"{md(a)} 결품"))
    if pers_s:
        a = pers_s[0]
        cands.append((max(a, lt) - lt, 1, f"{pre}안전재고 회복 불가" if a == 0 else f"{pre}{md(a)} 안전재고 이탈"))
    cands.sort()
    d_pol = cands[0][0] if cands else None
    reason = cands[0][2] if cands else ""

    gap_txt = dip_txt = ""
    if trans_h:
        a0, b0 = trans_h[0]
        tot = sum(b - a + 1 for a, b in trans_h)
        gap_txt = (f"입고 전 결품 {tot}일 ({md(a0)}~{md(b0)}" + (f", {md(b0 + 1)} 입고" if b0 + 1 < n else "") + ")"
                   + (f" 외 {len(trans_h) - 1}회" if len(trans_h) > 1 else ""))
    elif trans_s:
        a0, b0 = trans_s[0]
        tot = sum(b - a + 1 for a, b in trans_s)
        dip_txt = (f"입고 전 안전재고 이탈 {tot}일 ({md(a0)}~{md(b0)}" + (f", {md(b0 + 1)} 입고" if b0 + 1 < n else "") + ")")

    if not has_demand:
        tier, status = 5, "💤 출고없음"
    elif unavoid:
        a = unavoid[0]
        tier, status = 0, f"🚨 결품 불가피 {lt - a}일 ({md(a)}~{md(lt - 1)}) · 오늘 발주해도 신규분 입고({md(lt)}) 전 부족"
    elif d_pol is None:
        tier, status = 4, f"✅ 안정 ({n}일 내 이상 없음)"
    elif d_pol <= 0:
        tier, status = 1, f"🚨 오늘 발주 ({reason})"
    elif d_pol <= 7:
        tier, status = 2, f"⚠️ 긴급 D-{d_pol} ({reason})"
    elif d_pol <= 14:
        tier, status = 3, f"🔔 준비 D-{d_pol} ({reason})"
    else:
        tier, status = 4, f"✅ 여유 D-{d_pol}"

    if has_demand and not unavoid:
        if gap_txt:
            if tier >= 4 and trans_h[0][0] < lt:         # 신규 발주로는 못 막는 임박한 일시 결품 → 발주 대기가 아니라 입고/출고 일정 점검 대상
                tier = 1.5
                status = "🟠 " + gap_txt + ("" if d_pol is None else f" · 신규 발주는 D-{d_pol}")
            else:
                status += " · " + gap_txt
        elif dip_txt and tier >= 4:
            status += " · " + dip_txt

    # 권장 발주량: 신규 발주분이 도착하는 시점부터 target_days 동안 평택 재고가 기준선(안전재고, 없으면 0) 아래로 내려가는 최대 부족분
    qty = 0.0
    if has_demand and d_pol is not None and d_pol + lt < n:
        A = d_pol + lt
        floor = safety if safety > 0 else 0.0
        base_lost = sim["lost"][A - 1] if A >= 1 else 0.0
        end = min(n, A + target_days + 1)
        qty = max(0.0, max(floor - ms[i] + (sim["lost"][i] - base_lost) for i in range(A, end)))
        if not item.moq_all and item.moq > 0:
            qty = math.ceil(qty / item.moq - 1e-9) * item.moq
    return dict(tier=tier, status=status, d_pol=d_pol, d_stock=None,
                safety_day=(soft[0][0] if soft else None), reason=reason, gap=gap_txt, dip=dip_txt,
                deadline=(base + timedelta(days=d_pol)) if d_pol is not None else None, rec_qty=int(round(qty)))


def fmt_md(base: date, offset):
    return "-" if offset is None else (base + timedelta(days=offset)).strftime("%m/%d")


# ─────────────────────────────── 쿠팡 PO 예측 ───────────────────────────────
def norm_name(s) -> str:
    return re.sub(r"[\s,.\[\]()（）_+\-/·]", "", str(s)).lower()


def auto_match(po_names, item_names) -> dict:
    """PO 품명 → (재고시트 품명 또는 None, 매칭 방식). 공백·기호를 뺀 이름이 같으면 일치, 아니면 유사도 60% 이상 1순위."""
    by_norm = {}
    for n in item_names:
        by_norm.setdefault(norm_name(n), n)
    out, used = {}, set()
    for p in po_names:
        n = by_norm.get(norm_name(p))
        if n:
            out[p] = (n, "일치")
            used.add(n)
    rest = [n for n in item_names if n not in used]
    for p in po_names:
        if p in out:
            continue
        k = norm_name(p)
        scored = sorted(((difflib.SequenceMatcher(None, k, norm_name(n)).ratio(), n) for n in rest), reverse=True)
        if scored and scored[0][0] >= 0.6 and (len(scored) == 1 or scored[0][0] - scored[1][0] >= 0.05):
            out[p] = (scored[0][1], f"유사 {scored[0][0]:.0%}")
        else:
            out[p] = (None, "없음")
    return out


def _skip_weekend(d: date) -> date:
    wd = d.weekday()
    return d + timedelta(days=7 - wd) if wd >= 5 else d       # 토·일 → 다음 월요일


def forecast_po(name: str, recs: list, base: date, end: date, item, cfg: dict) -> dict:
    """
    PO 이력(한 품명) → 차기 PO 예측 + 이벤트 목록(실적/확정/예측).
      Q (PO 1회 수량) = 최근 N회 PO(일자별 합, 납품>0)의 중앙값
      r (일 출고율)   = VF형: 시트 최근 출고 평균 / 밀크런·미매칭: 최근 W일 PO 실적 ÷ W
      간격            = Q ÷ r
    """
    recs = sorted(recs, key=lambda x: x["date"])
    daily = {}
    for x in recs:
        daily[x["date"]] = daily.get(x["date"], 0.0) + x["qty"]
    pos = sorted((d, q) for d, q in daily.items() if q > 0)
    past = [(d, q) for d, q in pos if d <= base]
    fut = [(d, q) for d, q in pos if d > base]
    price = next((x["price"] for x in reversed(recs) if x["price"] > 0), 0.0)
    W = cfg["win_po"]
    r_po = sum(q for d, q in past if (base - d).days < W) / W
    zero60 = sum(1 for x in recs if x["qty"] <= 0 and 0 <= (base - x["date"]).days < 60)

    res = dict(name=name, item=item, price=price, Q=None, rate=0.0, r_po=r_po, r_sheet=None, rate_src="",
               gap=None, next=None, overdue=False, last=(past[-1][0] if past else None),
               fut_last=(fut[-1][0] if fut else None), zero60=zero60, note="", events=[])

    # 실적·확정 이벤트 (원래 행 단위, 단가는 행의 단가 우선)
    for x in recs:
        if x["qty"] <= 0:
            continue
        kind = "실적" if x["date"] <= base else "확정"
        res["events"].append(dict(date=x["date"], kind=kind, demand=x["qty"], deliv=x["qty"],
                                  price=x["price"] if x["price"] > 0 else price))
    if not pos:
        res["note"] = "납품 수량이 있는 PO 없음"
        return res

    q_hist = [q for _, q in (past or pos)][-cfg["q_n"]:]
    Q = float(statistics.median(q_hist))
    res["Q"] = Q

    use_sheet = cfg["mode"] == "sheet" and item is not None and not item.is_milkrun
    r_sheet = None
    if item is not None:
        days = [d for d in cfg["entered"] if d < base][-cfg["win_sheet"]:]
        r_sheet = sum(item.sales.get(d, 0.0) for d in days) / len(days) if days else 0.0
        res["r_sheet"] = r_sheet
    if use_sheet and r_sheet and r_sheet > 0:
        r, src = r_sheet, f"시트 출고 최근 {cfg['win_sheet']}일"
    elif use_sheet and r_po > 0:
        use_sheet = False
        r, src = r_po, f"PO 실적 {W}일 (최근 출고 0 → 대체)"
    else:
        use_sheet = False
        r, src = r_po, f"PO 실적 {W}일" + (" (밀크런)" if item is not None and item.is_milkrun else "")
    r *= cfg["mult"]
    res["rate"], res["rate_src"] = r, src
    if not use_sheet and len(pos) < 2:
        res["note"] = "PO 이력 1회뿐 → 간격 추정 불가, 예측 안 함"
        res["rate"] = 0.0
        return res
    if r <= 0:
        res["note"] = f"최근 {W}일 PO·출고 없음 → 예측 안 함"
        return res

    gap = Q / r
    res["gap"] = gap
    if fut:                                         # 이미 잡힌 미래 PO가 있으면 그 다음부터
        raw = (fut[-1][0] - base).days + gap
    elif use_sheet and past:                        # 마지막 PO 이후 실제 출고가 Q를 채우는 시점
        last = past[-1][0]
        consumed = sum(item.sales.get(d, 0.0) for d in cfg["entered"] if last < d < base)
        raw = (Q - consumed * cfg["mult"]) / r
    else:
        raw = (past[-1][0] - base).days + gap
    res["overdue"] = raw < 0 and not fut
    t = max(1.0, raw)                               # 기준일 당일 PO는 이미 입력된 것으로 보고 다음 날부터 예측

    preds = []
    while len(preds) < 500:
        d = base + timedelta(days=int(round(t)))
        if cfg["weekend"]:
            d = _skip_weekend(d)
        if d > end:
            break
        preds.append(d)
        t += gap
    for d in preds:
        res["events"].append(dict(date=d, kind="예측", demand=Q, deliv=Q, price=price))
    res["next"] = preds[0] if preds else None

    # 재고 제약: 평택 재고 + 발주완료 입고분 안에서만 납품 가능 (확정·예측 PO를 날짜순으로 차감)
    if cfg["cap"] and item is not None:
        stock = item.main
        inb = sorted((d, q) for d, q, _ in item.inbound)
        ii = 0
        for e in sorted((e for e in res["events"] if e["kind"] != "실적"), key=lambda e: (e["date"], e["kind"] == "예측")):
            while ii < len(inb) and inb[ii][0] <= e["date"]:
                stock += inb[ii][1]
                ii += 1
            e["deliv"] = min(e["demand"], max(stock, 0.0))
            stock -= e["deliv"]
    res["events"].sort(key=lambda e: e["date"])
    return res


def forecast_all(po_recs, items, mapping, base, end, cfg):
    by_name = {}
    for r in po_recs:
        by_name.setdefault(r["name"], []).append(r)
    item_by = {i.name: i for i in items}
    return [forecast_po(nm, recs, base, end, item_by.get(mapping.get(nm)), cfg) for nm, recs in by_name.items()]


def ev_rev(e, vat_incl: bool, demand: bool = False) -> float:
    return (e["demand"] if demand else e["deliv"]) * e["price"] * (1.0 if vat_incl else 1 / 1.1)


# ─────────────────────────────── UI ───────────────────────────────
_ST_VER = tuple(int(x) for x in re.findall(r"\d+", st.__version__)[:2])
_HAS_SELECT = _ST_VER >= (1, 35)       # st.dataframe(on_select=...) 지원 버전
SKU_KEY = "sku_sel"


def wide(fn, *args, **kw):
    """Streamlit 버전별 '가로 꽉 채우기' 인자 차이 흡수."""
    order = ({"width": "stretch"}, {"use_container_width": True}) if _ST_VER >= (1, 50) \
        else ({"use_container_width": True}, {"width": "stretch"})
    for extra in order:
        try:
            return fn(*args, **kw, **extra)
        except Exception:
            continue
    return fn(*args, **kw)


def num_col(label):
    """천 단위 쉼표 표시 (구버전 Streamlit은 정수 표시)."""
    return st.column_config.NumberColumn(label, format="localized" if _ST_VER >= (1, 39) else "%d")


def won(v: float) -> str:
    return f"{v / 1e4:,.0f}만원"


def render_po_tab(fcs, base, horizon_end, vat_incl, cfg):
    vat_txt = "VAT 포함" if vat_incl else "VAT 제외(공급가)"
    if not fcs:
        st.info("엑셀에 쿠팡 PO 이력 탭이 없습니다. 헤더에 '날짜'와 '단가'가 있는 탭을 넣으면 자동으로 인식합니다.")
        return
    all_dates = [e["date"] for f in fcs for e in f["events"]]
    first = min(all_dates) if all_dates else base
    months, (y, m) = [], (first.year, first.month)
    while (y, m) <= (horizon_end.year, horizon_end.month):
        months.append((y, m))
        y, m = add_months(y, m, 1)
    cur = (base.year, base.month)

    c1, c2 = st.columns([1, 2])
    with c1:
        sel = st.selectbox("📅 매출을 볼 월", months, index=months.index(cur) if cur in months else len(months) - 1,
                           format_func=lambda t: f"{t[0]}년 {t[1]}월")
    cats = sorted({(f["item"].cat if f["item"] else "미매칭") for f in fcs})
    with c2:
        sel_cats = st.multiselect("구분 필터 (비우면 전체)", cats, default=[])
    view = [f for f in fcs if not sel_cats or (f["item"].cat if f["item"] else "미매칭") in sel_cats]
    sy, sm = sel
    ly = (sy - 1, sm)
    mk = lambda d: (d.year, d.month)

    # ── KPI
    def tot(kind=None, month=sel, demand=False):
        return sum(ev_rev(e, vat_incl, demand) for f in view for e in f["events"]
                   if mk(e["date"]) == month and (kind is None or e["kind"] == kind))
    r_act, r_conf, r_pred = tot("실적"), tot("확정"), tot("예측")
    r_all = r_act + r_conf + r_pred
    r_dem = tot(None, demand=True)
    r_ly = tot("실적", ly)
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric(f"💰 {sy}년 {sm}월 예상 매출", won(r_all),
              delta=(f"전년 동월 대비 {(r_all / r_ly - 1) * 100:+.0f}%" if r_ly > 0 else None))
    k2.metric("✅ 실적 (기준일까지)", won(r_act))
    k3.metric("📌 확정 예정 PO", won(r_conf))
    k4.metric("🔮 예측 PO", won(r_pred))
    k5.metric("⛔ 재고 부족으로 빠진 매출", won(r_dem - r_all),
              help="재고 제약을 끄면 이 금액까지 포함한 '수요 기준' 매출이 됩니다.")
    st.caption(f"{vat_txt} · 기준일 {base:%Y-%m-%d} · 전년 동월({ly[0]}년 {ly[1]}월) 실적 {won(r_ly)}"
               + ("" if cfg["cap"] else " · 재고 제약 미반영(수요 기준)"))

    # ── 품목별 표
    rows = []
    for f in view:
        it = f["item"]
        ev = [e for e in f["events"] if mk(e["date"]) == sel]
        part = lambda k: [e for e in ev if e["kind"] == k]
        rev = lambda es, d=False: sum(ev_rev(e, vat_incl, d) for e in es)
        nxt = f["next"]
        rows.append({
            "구분": it.cat if it else "미매칭",
            "품명(PO)": f["name"],
            "유형": ("밀크런" if it.is_milkrun else "VF") if it else "-",
            "단가": int(f["price"]),
            "PO 1회 수량": int(f["Q"]) if f["Q"] else None,
            "일 출고율": round(f["rate"], 1) if f["rate"] else None,
            "출고율 근거": f["rate_src"] or f["note"],
            "PO 간격(일)": round(f["gap"], 1) if f["gap"] else None,
            "마지막 PO": f["last"].strftime("%m/%d") if f["last"] else "-",
            "확정 PO ~": f["fut_last"].strftime("%m/%d") if f["fut_last"] else "-",
            "차기 PO 예상": (nxt.strftime("%m/%d") + (" ⏰지연" if f["overdue"] else "")) if nxt else "-",
            "D-": (nxt - base).days if nxt else None,
            "선택월 PO 건수": len(ev),
            "선택월 수량": int(sum(e["deliv"] for e in ev)),
            "선택월 매출": int(rev(ev)),
            "└ 실적": int(rev(part("실적"))),
            "└ 확정": int(rev(part("확정"))),
            "└ 예측": int(rev(part("예측"))),
            "재고부족 차감": int(rev(ev, True) - rev(ev)),
            "전년 동월 매출": int(sum(ev_rev(e, vat_incl) for e in f["events"] if e["kind"] == "실적" and mk(e["date"]) == ly)),
            "0납품 PO(60일)": f["zero60"] or None,
        })
    df = pd.DataFrame(rows).sort_values(["선택월 매출", "D-"], ascending=[False, True]).reset_index(drop=True)
    st.subheader(f"📋 품목별 차기 PO & {sm}월 예상 매출")
    money = ["단가", "선택월 매출", "└ 실적", "└ 확정", "└ 예측", "재고부족 차감", "전년 동월 매출"]
    wide(st.dataframe, df, hide_index=True, column_config={c: num_col(c) for c in money + ["선택월 수량", "PO 1회 수량"]})
    st.caption("차기 PO 예상: VF형은 마지막 PO 이후 실제 출고 누적이 'PO 1회 수량'을 채우는 날, 밀크런은 마지막(확정 포함) PO + PO 간격. "
               "⏰지연 = 예상일이 이미 지났는데 PO가 없음 (재고 부족·쿠팡 발주 보류 등 확인 필요). "
               "재고부족 차감 = 수요상 PO는 들어오지만 평택 재고가 모자라 납품 못 할 것으로 보는 금액."
               + (" 토·일 예측분은 월요일로 옮깁니다." if cfg["weekend"] else ""))

    # ── 월별 매출 추이
    st.subheader("📈 월별 쿠팡 PO 매출 추이 (실적 + 확정 + 예측)")
    y0, m0 = add_months(cur[0], cur[1], -12)
    chart_months = [mm for mm in months if mm >= (y0, m0)]
    labels = [f"{a % 100:02d}.{b:02d}" for a, b in chart_months]
    agg = {k: [] for k in ("실적", "확정", "예측", "수요", "전년")}
    for mm in chart_months:
        agg["실적"].append(tot("실적", mm))
        agg["확정"].append(tot("확정", mm))
        agg["예측"].append(tot("예측", mm))
        agg["수요"].append(tot(None, mm, demand=True))
        agg["전년"].append(tot("실적", (mm[0] - 1, mm[1])))
    fig = go.Figure()
    for k, col in (("실적", "#1f4e79"), ("확정", "#2a9d8f"), ("예측", "#9ecae1")):
        fig.add_trace(go.Bar(x=labels, y=agg[k], name=k if k != "예측" else "예측(재고 반영)" if cfg["cap"] else "예측",
                             marker_color=col, hovertemplate="%{y:,.0f}원"))
    if cfg["cap"]:
        fig.add_trace(go.Scatter(x=labels, y=agg["수요"], name="수요 기준(재고 무제한)", mode="lines+markers",
                                 line=dict(color="#e76f51", dash="dash"), hovertemplate="%{y:,.0f}원"))
    fig.add_trace(go.Scatter(x=labels, y=agg["전년"], name="전년 동월 실적", mode="lines+markers",
                             line=dict(color="#888888", dash="dot"), hovertemplate="%{y:,.0f}원"))
    fig.update_layout(barmode="stack", hovermode="x unified", yaxis_title=f"매출 (원, {vat_txt})",
                      margin=dict(l=20, r=20, t=30, b=20), legend=dict(orientation="h", y=-0.15))
    wide(st.plotly_chart, fig)

    # ── 품목별 PO 타임라인
    st.subheader("🔍 품목별 PO 타임라인")
    names = list(df["품명(PO)"])
    if st.session_state.get("po_pick") not in names:      # 필터 변경으로 선택값이 목록에서 빠지면 첫 품목으로
        st.session_state["po_pick"] = names[0]
    pick = st.selectbox("품목 선택:", names, key="po_pick")
    f = next(x for x in view if x["name"] == pick)
    lo = base - timedelta(days=90)
    evs = [e for e in f["events"] if e["date"] >= lo]
    fig2 = go.Figure()
    for k, col in (("실적", "#1f4e79"), ("확정", "#2a9d8f"), ("예측", "#9ecae1")):
        es = [e for e in evs if e["kind"] == k]
        if es:
            fig2.add_trace(go.Bar(x=[e["date"] for e in es], y=[e["deliv"] for e in es], name=k, marker_color=col))
    short = [e for e in evs if e["demand"] - e["deliv"] > 1e-9]
    if short:
        fig2.add_trace(go.Bar(x=[e["date"] for e in short], y=[e["demand"] - e["deliv"] for e in short],
                              name="재고 부족(미납 예상)", marker_color="#e76f51", opacity=0.6))
    fig2.add_vline(x=base.isoformat(), line_dash="dash", line_color="#999")
    fig2.update_layout(barmode="stack", yaxis_title="PO 수량", margin=dict(l=20, r=20, t=30, b=20),
                       legend=dict(orientation="h", y=-0.15))
    wide(st.plotly_chart, fig2)
    it = f["item"]
    st.caption(f"PO 1회 {f['Q'] or 0:,.0f}개 · 일 출고율 {f['rate']:,.1f}개 ({f['rate_src'] or f['note']}) · "
               f"참고: PO 실적 기준 {f['r_po']:,.1f}개/일"
               + (f", 시트 출고 기준 {f['r_sheet']:,.1f}개/일" if f["r_sheet"] is not None else "")
               + (f" · 매칭 품목: {it.name} (평택 {it.main:,.0f}개)" if it else " · 재고시트 매칭 없음 → 재고 제약 미적용"))


def main():
    st.set_page_config(layout="wide", page_title="쿠팡 VF & 평택 마스터 발주 대시보드")
    st.title("📦 쿠팡 VF & 평택물류 통합 마스터 발주 시스템")
    st.caption("월별 시트 자동 연동 · 실적 기반 리드타임 · 쿠팡 PO 차감 시뮬레이션 · 안전재고/결품 기준 공장 발주 데드라인 · 쿠팡 PO 예측/월 매출")

    st.sidebar.header("📁 엑셀 파일 업로드")
    up = st.sidebar.file_uploader("통합 마스터 엑셀 업로드 (.xlsx)", type=["xlsx"])
    if not up:
        st.info("👈 왼쪽 사이드바에서 마스터 엑셀 파일을 업로드해 주세요.")
        return
    raw_bytes = up.getvalue()
    sig = hashlib.md5(raw_bytes).hexdigest()[:8]
    try:
        wb = load_workbook(raw_bytes, PARSER_VERSION)
    except Exception as e:
        st.error("엑셀을 읽는 중 오류가 발생했습니다.")
        st.exception(e)
        return
    sheets = [s for s in wb["sheets"] if s["rows"]]
    po_recs = [r for p in wb["po"] for r in p["rows"]]
    if wb["skipped"]:
        st.warning("형식을 인식하지 못해 건너뛴 시트 → " + " / ".join(wb["skipped"]))
    if not sheets:
        st.error("품목 행을 찾지 못했습니다. 헤더에 '품명' 컬럼이 있는지 확인해 주세요.")
        return
    bad = [f"{s['name']}: {', '.join(s['missing'])}" for s in sheets if s["missing"]]
    if bad:
        st.warning("일부 컬럼을 인식하지 못했습니다 (해당 값은 0으로 처리) → " + " / ".join(bad))

    snaps = [s["snapshot"] for s in sheets if s["snapshot"]]
    snapshot = max(snaps) if snaps else None

    st.sidebar.divider()
    st.sidebar.header("⚙️ 발주 및 시뮬레이션 기본값")
    base = st.sidebar.date_input(
        "재고/출고 기준일자", snapshot or today_kst(), key=f"base_{sig}",
        help="재고 수량이 입력된 날짜입니다. 기본값은 엑셀 헤더의 기준일(=TODAY() 저장값), 없으면 오늘(KST).")
    if snapshot and abs((today_kst() - snapshot).days) >= 2:
        st.sidebar.warning(f"엑셀 기준일은 {snapshot:%m/%d}, 오늘은 {today_kst():%m/%d} 입니다. 재고가 최신인지 확인하세요.")
    win_std = st.sidebar.number_input("출고 평균 기간 – 일반(벤플 등, 일)", 1, 60, 7)
    win_mr = st.sidebar.number_input("출고 평균 기간 – 밀크런(일)", 1, 60, 28,
                                     help="밀크런은 PO 단위로 몰아서 나가서 7일 평균은 0이 되기 쉽습니다.")
    mr_kw = st.sidebar.text_input("밀크런 구분 키워드", "밀크런")
    default_lt = st.sidebar.number_input("기본 리드타임 (일)", 1, 120, 35)
    use_hist_lt = st.sidebar.checkbox("발주일→입고일 이력으로 리드타임 자동 추정", True,
                                      help="엑셀의 '발주일/입고일' 쌍에서 품목군별 중앙값을 계산해 적용합니다.")
    trig_mode_label = st.sidebar.radio("쿠팡 PO 트리거 방식", ["VF 잔여일수 기준", "고정 수량 기준"])
    trig_mode = "days" if trig_mode_label.startswith("VF") else "units"
    if trig_mode == "days":
        trig_value = st.sidebar.number_input("VF 잔여 며칠분 밑이면 PO?", 1, 60, 10,
                                             help="가정값입니다. 실제 쿠팡 PO 발생 이력으로 보정하세요.")
    else:
        trig_value = st.sidebar.number_input("VF 잔여 수량 트리거 (개)", 0, 100000, 200, step=10)
    po_mult = st.sidebar.selectbox("쿠팡 PO 배수", [1, 2, 3], format_func=lambda x: f"MOQ × {x}")
    transfer_days = st.sidebar.number_input("평택→VF 이동 소요일", 0, 14, 0)
    target_days = st.sidebar.number_input("권장 발주량 산정 기간 (일)", 7, 120, 45,
                                          help="신규 발주분이 도착한 뒤 이 기간 동안 부족이 없도록 산정합니다.")
    use_plan = st.sidebar.checkbox("시트에 미리 입력된 미래 출고(쿠팡 PO 예정) 반영", True,
                                   help="기준일 이후 날짜 칸에 이미 수량이 적혀 있으면 그 기간은 평균 대신 그 값을 그대로 출고로 사용합니다.")
    include_stale = st.sidebar.checkbox("최신 월 시트에 없는 품목도 포함", False)

    st.sidebar.divider()
    st.sidebar.header("🛒 쿠팡 PO 예측 · 매출")
    po_mode_label = st.sidebar.radio("출고율 기준", ["VF형=시트 최근 출고 / 밀크런=PO 실적", "모두 PO 실적 기준"],
                                     help="VF형 PO는 고객 출고만큼 VF를 채우러 들어오므로 최근 출고량이 차기 PO를 가장 잘 설명합니다. "
                                          "밀크런은 시트 출고 = PO 납품이라 PO 실적을 씁니다.")
    po_win_sheet = st.sidebar.number_input("시트 출고 평균 기간 (일)", 1, 60, 7)
    po_win_po = st.sidebar.number_input("PO 실적 집계 기간 (일)", 7, 120, 28)
    po_q_n = st.sidebar.number_input("PO 1회 수량 = 최근 N회 중앙값, N", 1, 30, 8)
    po_mult_pct = st.sidebar.number_input("수요 보정 (%)", 10, 500, 100, step=5,
                                          help="시즌·행사 등으로 앞으로 출고가 늘거나 줄 것으로 보면 조정하세요. 예측 PO 간격이 그만큼 짧아/길어집니다.")
    po_cap = st.sidebar.checkbox("재고 제약 반영 (평택 재고 + 발주완료 입고분 안에서만 납품)", True)
    po_weekend = st.sidebar.checkbox("토·일 예측 PO는 월요일로", True, help="PO 이력상 일요일 PO는 없고 토요일도 드뭅니다.")
    vat_incl = st.sidebar.checkbox("매출 VAT 포함", True, help="PO 탭 단가가 VAT 포함 기준입니다. 끄면 ÷1.1 공급가로 봅니다.")

    # ── 입고 일정 보정: '11월초' 같은 추정 표기를 실제 날짜로 직접 고칠 수 있음 (최신 월 시트 기준)
    latest_sheet = sorted(sheets, key=lambda s_: (s_["year"], s_["month"]))[-1]
    overrides = {}
    if latest_sheet["inbound_cols"]:
        n_items = {}
        for r_ in latest_sheet["rows"]:
            for (*_, col_) in r_["inbound"]:
                n_items[col_] = n_items.get(col_, 0) + 1
        icols = latest_sheet["inbound_cols"]
        tbl = pd.DataFrame(
            [{"품목군": c["group"], "원본 발주일": c.get("raw_order") or "-", "원본 입고일": c.get("raw_arrival", ""),
              "수량 있는 품목": n_items.get(c["col"], 0), "발주일": c["order"], "입고일": c["arrival"]} for c in icols],
            index=[c["col"] for c in icols])
        with st.expander("🗓️ 입고 일정 보정 (발주일·입고일 직접 수정)", expanded=False):
            st.caption("엑셀에 '11월초' 처럼 대략 적힌 날짜를 실제 날짜로 바꾸면 그 날짜가 확정으로 계산에 쓰입니다. "
                       "발주일이 기준일 당일 또는 이전이면 발주 완료로 반영됩니다. 새 파일을 올리면 보정은 초기화되니, 계속 쓰려면 엑셀 헤더에 날짜를 적어 두세요.")
            edited = st.data_editor(
                tbl, hide_index=True, key=f"inb_edit_{sig}",
                disabled=["품목군", "원본 발주일", "원본 입고일", "수량 있는 품목"],
                column_config={"발주일": st.column_config.DateColumn("발주일", format="YYYY-MM-DD"),
                               "입고일": st.column_config.DateColumn("입고일", format="YYYY-MM-DD")})
        for col_idx, row_ in edited.iterrows():
            o_new, a_new = as_date(row_["발주일"]), as_date(row_["입고일"])
            orig = next(c for c in icols if c["col"] == col_idx)
            if a_new is not None and (o_new, a_new) != (orig["order"], orig["arrival"]):
                overrides[col_idx] = (o_new, a_new)
    sheets = apply_inbound_overrides(sheets, latest_sheet["name"], overrides)

    items, latest, entered, lt_est = build_items(
        sheets, base, win_std, win_mr, default_lt, use_hist_lt, mr_kw, include_stale, use_plan)
    if not items:
        st.error("표시할 품목이 없습니다.")
        return
    excluded = sorted({r["name"] for s in sheets for r in s["rows"]} - {i.name for i in items})

    # ── PO 품명 ↔ 재고시트 품명 매칭 (자동 + 수동 보정)
    po_names = sorted({r["name"] for r in po_recs})
    auto = auto_match(po_names, [i.name for i in items])
    mapping = {p: v[0] for p, v in auto.items() if v[0]}
    if po_names:
        with st.expander("🔗 PO 품명 ↔ 재고시트 품명 매칭", expanded=any(v[0] is None or v[1] != "일치" for v in auto.values())):
            st.caption("이름이 달라 자동으로 '유사' 매칭했거나 못 찾은 품목은 여기서 직접 고르세요. "
                       "매칭돼야 시트 출고율·평택 재고가 PO 예측에 쓰입니다. 매칭이 없으면 PO 실적만으로 예측하고 재고 제약은 빠집니다.")
            mtbl = pd.DataFrame([{"PO 품명": p, "자동 매칭": auto[p][1], "재고시트 품명": auto[p][0] or NO_MATCH}
                                 for p in po_names])
            medit = st.data_editor(
                mtbl, hide_index=True, key=f"po_map_{sig}", disabled=["PO 품명", "자동 매칭"],
                column_config={"재고시트 품명": st.column_config.SelectboxColumn(
                    "재고시트 품명", options=[NO_MATCH] + sorted(i.name for i in items), required=True)})
        mapping = {r_["PO 품명"]: r_["재고시트 품명"] for _, r_ in medit.iterrows() if r_["재고시트 품명"] != NO_MATCH}

    hy, hm = add_months(base.year, base.month, FORECAST_MONTHS - 1)
    horizon_end = month_end(hy, hm)
    po_cfg = dict(mode="sheet" if po_mode_label.startswith("VF") else "po", win_sheet=int(po_win_sheet),
                  win_po=int(po_win_po), q_n=int(po_q_n), mult=po_mult_pct / 100, cap=po_cap,
                  weekend=po_weekend, entered=entered)
    fcs = forecast_all(po_recs, items, mapping, base, horizon_end, po_cfg) if po_recs else []
    next_po = {}
    for f in fcs:
        if f["item"] is not None and f["next"]:
            n0 = next_po.get(f["item"].name)
            next_po[f["item"].name] = min(n0, f["next"]) if n0 else f["next"]

    # ── 전 품목 계산
    calc = {}
    for it in items:
        sim = simulate(it, base, it.adu, trig_mode, trig_value, po_mult, transfer_days)
        calc[it.name] = (sim, evaluate(it, sim, base, it.adu, it.safety, it.lt, target_days))

    rows = []
    for it in items:
        sim, ev = calc[it.name]
        total = it.main + it.vf
        rows.append({
            "_tier": ev["tier"], "_dpol": ev["d_pol"] if ev["d_pol"] is not None else 9999,
            "구분": it.cat, "품명": it.name, "유형": "밀크런" if it.is_milkrun else "VF",
            "평택재고": int(it.main), "VF재고": None if it.is_milkrun else int(it.vf),
            "안전재고": int(it.safety) if it.safety > 0 else None,
            "안전재고 대비(%)": round(it.main / it.safety * 100) if it.safety > 0 else None,
            "일평균출고": round(it.adu, 1), "재고커버(일)": round(total / it.adu) if it.adu > 0 else None,
            "적용L/T(일)": it.lt, "L/T 근거": it.lt_src,
            "차기 PO(이력예측)": next_po[it.name].strftime("%m/%d") if it.name in next_po else "-",
            "차기 쿠팡PO(시뮬)": fmt_md(base, (sim["first_po"] - base).days) if sim["first_po"] else "-",
            "안전재고 이탈일": fmt_md(base, ev["safety_day"]), "평택 결품일": fmt_md(base, sim["stockout"]),
            "발주 데드라인": ev["deadline"].strftime("%m/%d") if ev["deadline"] else "-",
            "발주상태": ev["status"], "권장 발주량": ev["rec_qty"] if it.adu > 0 else None,
            "입고예정(발주완료)": int(sum(q for _, q, _ in it.inbound)) or None,
            "입고계획(미발주)": int(sum(x["qty"] for x in it.inbound_detail if not x["ordered"])) or None,
            "출고예정(합)": int(sum(it.plan.values())) or None,
        })
    df = pd.DataFrame(rows)

    tab1, tab2 = st.tabs(["📦 재고 · 공장 발주", "🛒 쿠팡 PO 예측 · 월 매출"])
    with tab2:
        render_po_tab(fcs, base, horizon_end, vat_incl, po_cfg)

    with tab1:
        cats = ["전체 보기"] + sorted(df["구분"].unique())
        sel_cat = st.selectbox("📂 구분(카테고리) 필터:", cats)
        view = df if sel_cat == "전체 보기" else df[df["구분"] == sel_cat]
        view = view.sort_values(["_tier", "_dpol"]).reset_index(drop=True)

        std_days = [d for d in entered if d < base][-win_std:]
        mr_days = [d for d in entered if d < base][-win_mr:]
        period = f"{std_days[0]:%m/%d}~{std_days[-1]:%m/%d} ({len(std_days)}일)" if std_days else "날짜 컬럼 없음"
        if mr_days:
            period += f" · 밀크런 {len(mr_days)}일"
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("📅 기준일자", f"{base:%Y-%m-%d}")
        k2.metric("📊 출고 평균 구간", period)
        k3.metric("🚨 즉시/지연 발주", f"{int((view['_tier'] <= 1).sum())} 개 SKU")
        k4.metric("⚠️ 14일 내 발주", f"{int(view['_tier'].isin([2, 3]).sum())} 개 SKU")

        notes = []
        if overrides:
            notes.append(f"입고 일정 {len(overrides)}개 컬럼을 화면에서 직접 보정했습니다 (보정한 날짜는 확정으로 계산).")
        not_yet = [(it, x) for it in items for x in it.inbound_detail if not x["ordered"]]
        if not_yet:
            cols_ = {(x["group"], x["order"], x["arrival"]) for _, x in not_yet}
            notes.append(f"발주일이 기준일({base:%m/%d})보다 뒤인 입고 계획 {len(cols_)}개 컬럼(품목 {len({it.name for it, _ in not_yet})}개)은 "
                         "아직 발주 전으로 보고 재고 계산에서 제외했습니다. 표의 '입고계획(미발주)'에 합계만 참고용으로 표시합니다.")
        if any(x["order"] is None and x["arrival"] >= base for it in items for x in it.inbound_detail):
            notes.append("발주일이 적혀 있지 않은 입고 컬럼은 발주 완료로 간주했습니다.")
        est_cnt = sum(1 for it in items for _, _, e in it.inbound if e)
        if est_cnt:
            notes.append(f"입고일이 '11월초/12월 중순' 같은 표현인 {est_cnt}건은 순의 끝날(초=10일, 중순=20일, 말=말일)로 보수적으로 추정했습니다.")
        for sh in sheets:
            if sh["unparsed"]:
                notes.append(f"[{sh['name']}] 날짜를 해석하지 못한 입고 컬럼: {', '.join(sh['unparsed'])}")
        if excluded:
            notes.append(f"최신 시트({latest})에 없어 제외한 품목 {len(excluded)}개: {', '.join(excluded)}")
        planned = [i for i in items if i.plan]
        if planned:
            last = max(i.plan_until for i in planned)
            notes.append(f"기준일 이후 날짜에 이미 입력된 출고 수량({len(planned)}개 품목, ~{last:%m/%d})은 쿠팡 PO 예정으로 보고 "
                         "그 기간의 출고로 그대로 사용했습니다. 예정이 아니라면 사이드바에서 반영을 끄세요.")
        if any(i.moq_all for i in items):
            notes.append("납품 MOQ가 '전량'인 품목은 PO 시 평택 재고 전부를 VF로 보내는 것으로 계산합니다.")
        if any(i.moq_missing for i in items):
            notes.append(f"MOQ가 비어 있는 품목은 임시로 {DEFAULT_MOQ}개로 계산했습니다.")
        for p in wb["po"]:
            ds = [r["date"] for r in p["rows"]]
            if ds:
                notes.append(f"[{p['name']}] 쿠팡 PO {len(ds)}행 인식 ({min(ds):%Y-%m-%d} ~ {max(ds):%Y-%m-%d}), "
                             f"기준일 이후 확정 PO {sum(1 for d in ds if d > base)}행"
                             + (f", 날짜를 못 읽은 행 {p['bad_date']}개 제외" if p["bad_date"] else ""))
        dup = pd.Series([(r["date"], r["name"], r["qty"]) for r in po_recs]).value_counts() if po_recs else pd.Series(dtype=int)
        dup = dup[dup > 1]
        if len(dup):
            notes.append(f"PO 탭에 날짜·품명·수량이 똑같은 행이 {len(dup)}묶음 있습니다(예: {dup.index[0][0]:%m/%d} {dup.index[0][1]}). "
                         "같은 날 PO가 2건이면 정상이고, 중복 입력이면 매출이 이중으로 잡히니 확인하세요.")
        unmatched = [p for p in po_names if p not in mapping]
        if unmatched:
            notes.append(f"재고시트와 매칭되지 않은 PO 품명 {len(unmatched)}개: {', '.join(unmatched)} (PO 실적만으로 예측)")
        no_po = sorted({i.name for i in items} - set(mapping.values()))
        if po_names and no_po:
            notes.append(f"PO 이력이 없는 재고시트 품목 {len(no_po)}개는 매출 예측에서 빠집니다: {', '.join(no_po)}")
        if notes:
            with st.expander(f"🔎 데이터 점검 사항 ({len(notes)})", expanded=False):
                for n in notes:
                    st.write("• " + n)
                if lt_est:
                    st.write("**발주→입고 실적 리드타임 (품목군별 중앙값)**")
                    st.dataframe(pd.DataFrame([{"품목군": g, "중앙값(일)": v["median"], "표본수": v["n"], "정확한 날짜 쌍": v["n_exact"]}
                                               for g, v in lt_est.items()]), hide_index=True)

        st.subheader("📋 전체 품목 발주 데드라인 마스터 테이블")
        show = view.drop(columns=["_tier", "_dpol", "L/T 근거"]).reset_index(drop=True)
        names = list(view["품명"])
        if _HAS_SELECT:
            event = wide(st.dataframe, show, hide_index=True, on_select="rerun", selection_mode="single-row", key="master_tbl")
            rows_sel = list(event.selection.rows) if event is not None and event.selection is not None else []
            picked = str(show.iloc[rows_sel[0]]["품명"]) if rows_sel else None
            if picked and picked != st.session_state.get("_last_row_pick"):
                st.session_state["_last_row_pick"] = picked        # 표에서 새로 고른 행만 아래 선택값을 덮어씀
                st.session_state[SKU_KEY] = picked
            elif not picked:
                st.session_state["_last_row_pick"] = None
            st.caption("👆 행 왼쪽 선택 칸을 클릭하면 아래 '개별 SKU' 차트가 그 품목으로 바뀝니다."
                       + (f"  현재 선택: **{st.session_state.get(SKU_KEY, names[0])}**" if names else ""))
        else:
            wide(st.dataframe, show, hide_index=True)
            st.caption("표 클릭 선택은 Streamlit 1.35 이상에서 지원됩니다. 아래 드롭다운으로 품목을 고르세요.")
        st.caption("발주 데드라인 = (발주완료 입고분을 반영하고도 끝까지 회복되지 않는 안전재고 이탈일/결품일) − 리드타임. "
                   "발주완료 입고분으로 곧 회복되는 일시 부족은 데드라인에 쓰지 않고 상태 칸에 '입고 전 결품/안전재고 이탈'로 따로 표시합니다(🟠 = 신규 발주가 아니라 입고·출고 일정 점검 대상). "
                   "권장 발주량은 신규분이 도착한 뒤 산정 기간 동안 안전재고(미기재 품목은 결품 방지) 기준 최대 부족분을 납품 MOQ 배수로 올림한 값입니다. "
                   "'차기 PO(이력예측)'은 쿠팡 PO 탭 이력 기반, '차기 쿠팡PO(시뮬)'은 사이드바 트리거 가정 기반입니다.")

        # ── 개별 SKU 시뮬레이터
        st.divider()
        st.subheader("🔍 개별 SKU 쿠팡 PO 차감 타임라인 & 재고 곡선")
        if st.session_state.get(SKU_KEY) not in names:          # 카테고리 필터 변경 등으로 선택값이 목록에 없으면 첫 품목으로
            st.session_state[SKU_KEY] = names[0]
        sku = st.selectbox("정밀 조회할 품목:", names, key=SKU_KEY)
        it = next(i for i in items if i.name == sku)
        h = hashlib.md5(sku.encode()).hexdigest()[:6]
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1:
            if trig_mode == "days":
                p_trig = st.number_input("PO 트리거 (VF 잔여 일수)", 1, 60, int(trig_value), key=f"tg_{h}")
            else:
                p_trig = st.number_input("PO 트리거 (VF 잔여 수량)", 0, 100000, int(trig_value), step=10, key=f"tg_{h}")
        with c2:
            p_mult = st.selectbox("쿠팡 PO 배수", [1, 2, 3], index=[1, 2, 3].index(po_mult), key=f"mu_{h}",
                                  format_func=lambda x: "전량" if it.moq_all else f"{x}배수 ({it.moq * x:,.0f}개)")
        with c3:
            p_adu = st.number_input("일일 출고량 (개/일)", 0.0, 1e6, float(round(it.adu, 2)), step=1.0, key=f"adu_{h}")
        with c4:
            p_safe = st.number_input("안전재고 (0이면 결품 기준)", 0, 10_000_000, int(it.safety), step=100, key=f"sf_{h}")
        with c5:
            p_lt = st.number_input("리드타임 (일)", 1, 120, int(it.lt), key=f"lt_{h}")
        p_tr = st.number_input("평택→VF 이동 소요일", 0, 14, int(transfer_days), key=f"tr_{h}") if not it.is_milkrun else 0
        st.caption(f"L/T 근거: {it.lt_src} · 출고평균 {it.adu_days}일 기준"
                   + (f" · {it.plan_until:%m/%d}까지는 시트 입력 예정 출고 사용" if it.plan_until else "")
                   + (f" · 시트 안전재고 환산 일평균 {it.safety / 45:,.1f}개/일 (안전재고÷45)" if it.safety > 0 else ""))

        sim = simulate(it, base, p_adu, trig_mode, p_trig, p_mult, p_tr, log=True)
        ev = evaluate(it, sim, base, p_adu, p_safe, p_lt, target_days)
        st.markdown(f"**발주상태:** {ev['status']}")
        m2, m3, m4 = st.columns(3)
        m2.metric("공장발주 데드라인", ev["deadline"].strftime("%Y-%m-%d") if ev["deadline"] else "-")
        m3.metric("평택 결품 예상일", fmt_md(base, sim["stockout"]))
        m4.metric("권장 발주량", f"{ev['rec_qty']:,}개" if p_adu > 0 else "-")

        days = [base + timedelta(days=i) for i in range(CHART_DAYS)]
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=days, y=sim["main"][:CHART_DAYS], mode="lines", name="평택본창고",
                                 line=dict(color="#1f77b4", width=3)))
        if not it.is_milkrun:
            fig.add_trace(go.Scatter(x=days, y=sim["vf"][:CHART_DAYS], mode="lines", name="VF 잔여재고",
                                     line=dict(color="#ff7f0e", width=2, dash="dash")))
        arr = [(d, q) for d, q, _ in it.inbound if base <= d < days[-1]]
        if arr:
            fig.add_trace(go.Scatter(x=[d for d, _ in arr], y=[sim["main"][(d - base).days] for d, _ in arr],
                                     mode="markers", name="공장 입고 (발주완료·반영)", marker=dict(symbol="triangle-up", size=12, color="#2ca02c")))
        plan_arr = [x for x in it.inbound_detail if not x["ordered"] and base <= x["arrival"] < days[-1]]
        if plan_arr:
            fig.add_trace(go.Scatter(x=[x["arrival"] for x in plan_arr], y=[sim["main"][(x["arrival"] - base).days] for x in plan_arr],
                                     mode="markers", name="입고 계획 (미발주·미반영)",
                                     marker=dict(symbol="triangle-up-open", size=12, color="#888888")))
        if p_safe > 0:
            fig.add_hline(y=p_safe, line_dash="dashdot", line_color="#ffbb00", annotation_text=f"안전재고 ({p_safe:,})")
        fig.update_layout(xaxis_title="일자", yaxis_title="수량 (개)", hovermode="x unified",
                          margin=dict(l=20, r=20, t=30, b=20))
        st.markdown(f"#### 📈 [{sku}] 향후 {CHART_DAYS}일 재고 흐름 (리드타임 {p_lt}일)")
        wide(st.plotly_chart, fig)

        with st.expander("📅 일자별 이벤트 로그 / 입고 스케줄"):
            if sim["events"]:
                wide(st.dataframe, pd.DataFrame(sim["events"], columns=["일자", "내용", "평택 잔여", "VF 잔여"]).round(0),
                     hide_index=True)
            else:
                st.write("기간 내 이벤트가 없습니다.")
            if it.inbound_detail:
                st.write("**공장 입고 스케줄** (발주일이 기준일 당일 또는 이전이면 발주완료로 반영)")
                st.dataframe(pd.DataFrame([{
                    "품목군": x["group"], "발주일": x["order"], "입고일": x["arrival"], "수량": int(x["qty"]),
                    "날짜": "추정" if x["est"] else "확정",
                    "상태": "✅ 발주완료 (반영)" if x["ordered"] else "🕓 미발주 계획 (미반영)"} for x in it.inbound_detail]),
                    hide_index=True)


if __name__ == "__main__":
    main()
