"""
쿠팡 VF & 평택(본창고) 통합 마스터 발주 · 매출 대시보드  (v4)

화면 구성
  🏠 홈            : 이번 달 매출·목표 진행률, 공장 발주 할 일, 입고·납품 점검, 7일 내 쿠팡 PO
  💰 매출 · 쿠팡 PO : 상품군별/품목별 월 매출(실적·확정·예측), 재고부족 손실, 월별 추이, SKU 상세
  📦 재고 · 공장 발주: 발주 데드라인 마스터 표, 개별 SKU 재고 곡선
  🏭 발주 시뮬레이터: 발주 수량·입고일을 넣어 결품·회복 매출 비교
  🛠️ 데이터 설정    : 입고 일정 보정, PO 품명 매칭, 상품군 규칙, 월 목표, 데이터 점검
  사이드바          : 엑셀 업로드, 기준일, VAT·재고 제약·수요 보정, 엑셀 리포트 다운로드, 고급 설정

선택 시트 '목표': [상품군 | 2026-10 | 2026-11 ...] 형태, 단위 만원 ('전체' 행 가능)

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
import dataclasses
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
# 쿠팡 PO 데이터에는 없지만 목표는 관리하는 상품군 (목표 표 맨 위에 표시, '기타매출' 시트 기본 상품군)
EXTRA_TARGET_GROUPS = ["네이버 두부모래"]


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


def parse_month_label(v, ref_year: int):
    if isinstance(v, (datetime, date)):
        return v.year, v.month
    t = clean_txt(v)
    for pat, f in ((r"^(\d{4})[-./년](\d{1,2})", lambda m: (int(m.group(1)), int(m.group(2)))),
                   (r"^(\d{2})년(\d{1,2})월", lambda m: (2000 + int(m.group(1)), int(m.group(2)))),
                   (r"^(\d{1,2})월$", lambda m: (ref_year, int(m.group(1))))):
        m = re.match(pat, t)
        if m and 1 <= f(m)[1] <= 12:
            return f(m)
    return None


def parse_target_sheet(raw: pd.DataFrame, ref_year: int) -> dict:
    """'목표' 시트: [상품군 | 2026-10 | 2026-11 ...] 형태, 단위 만원. → {(연, 월): {상품군: 만원}}"""
    n = raw.shape[0]
    hr = next((r for r in range(min(10, n)) if "상품군" in {clean_txt(x) for x in raw.iloc[r].tolist()}), None)
    if hr is None:
        raise ValueError("목표 시트에서 '상품군' 헤더를 찾지 못했습니다.")
    hdr = raw.iloc[hr].tolist()
    gcol = next(i for i, h in enumerate(hdr) if clean_txt(h) == "상품군")
    mcols = {c: parse_month_label(h, ref_year) for c, h in enumerate(hdr) if c != gcol}
    mcols = {c: ym for c, ym in mcols.items() if ym}
    out = {}
    for r in range(hr + 1, n):
        g = raw.iat[r, gcol]
        if not clean_txt(g):
            continue
        for c, ym in mcols.items():
            v = to_num(raw.iat[r, c])
            if v:
                out.setdefault(ym, {})[str(g).strip()] = v / 1e4 if v >= 1e6 else v   # 100만 이상이면 '원'으로 적은 것으로 보고 환산
    return out


def parse_manual_sales(raw: pd.DataFrame, ref_year: int) -> list:
    """'기타매출' 시트: [날짜 | 상품군 | 매출금액 | (메모)]. 날짜 대신 '2026-10' 같은 월만 적으면 그 달 1일로 봅니다."""
    n = raw.shape[0]
    hr = next((r for r in range(min(10, n))
               if any("매출" in clean_txt(x) or "금액" in clean_txt(x) for x in raw.iloc[r].tolist())
               and any(clean_txt(x) in ("날짜", "일자", "월", "기간") for x in raw.iloc[r].tolist())), None)
    if hr is None:
        raise ValueError("기타매출 시트에서 '날짜'와 '매출금액' 헤더를 찾지 못했습니다.")
    hdr = [clean_txt(x) for x in raw.iloc[hr].tolist()]
    c_d = next(i for i, h in enumerate(hdr) if h in ("날짜", "일자", "월", "기간"))
    c_a = next(i for i, h in enumerate(hdr) if "매출" in h or "금액" in h)
    c_g = next((i for i, h in enumerate(hdr) if h in ("상품군", "품명", "채널", "구분")), None)
    out = []
    for r in range(hr + 1, n):
        v = raw.iat[r, c_d]
        d = to_date(v)
        if d is None:
            ym = parse_month_label(v, ref_year)
            d = date(ym[0], ym[1], 1) if ym else None
        amt = to_num(raw.iat[r, c_a])
        if d is None or not amt:
            continue
        g = raw.iat[r, c_g] if c_g is not None else None
        out.append(dict(date=d, amount=amt, group=str(g).strip() if clean_txt(g) else EXTRA_TARGET_GROUPS[0]))
    return out


PARSER_VERSION = 6     # 파싱 결과 구조를 바꿀 때마다 올리면 Streamlit이 예전 캐시를 버리고 다시 파싱함


@st.cache_data(show_spinner="엑셀 분석 중...")
def load_workbook(file_bytes: bytes, parser_version: int = PARSER_VERSION) -> dict:
    """월별 재고 시트와 쿠팡 PO 이력 시트를 나눠 읽는다. 형식이 맞지 않는 시트는 건너뛰고 이유를 남긴다."""
    import io
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    sheets, po, skipped, targets, manual = [], [], [], {}, []
    for s in xl.sheet_names:
        raw = xl.parse(s, header=None)
        try:
            if "기타매출" in s or "매출입력" in s:
                manual += parse_manual_sales(raw, today_kst().year)
            elif "목표" in s:
                targets.update(parse_target_sheet(raw, today_kst().year))
            elif is_po_sheet(raw):
                po.append(parse_po_sheet(raw, s))
            else:
                sheets.append(parse_sheet(raw, s, today_kst().year))
        except Exception as e:          # 메모 탭 등 형식이 다른 시트 때문에 전체가 멈추지 않도록
            skipped.append(f"{s}: {e}")
    return dict(sheets=sheets, po=po, skipped=skipped, targets=targets, manual=manual)


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
               fut_last=(fut[-1][0] if fut else None), zero60=zero60, note="", events=[],
               recs=recs, q_hist=[], consumed=None, start_stock=None)

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
    res["Q"], res["q_hist"] = Q, q_hist

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
        res["consumed"] = consumed
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
        res["start_stock"] = stock
        inb = sorted((d, q) for d, q, _ in item.inbound)
        ii = 0
        for e in sorted((e for e in res["events"] if e["kind"] != "실적"), key=lambda e: (e["date"], e["kind"] == "예측")):
            while ii < len(inb) and inb[ii][0] <= e["date"]:
                stock += inb[ii][1]
                ii += 1
            avail = max(stock, 0.0)
            if e["kind"] == "확정":         # 이미 받은 PO는 매출에 그대로 넣고, 재고가 모자라면 '납품 위험'으로만 표시
                e["risk"] = max(0.0, e["demand"] - avail)
                stock -= e["demand"]        # 모자란 만큼은 이후 입고분에서 먼저 메워야 하는 것으로 봄
            else:
                e["deliv"] = min(e["demand"], avail)
                stock -= e["deliv"]
            e["stock_after"] = stock
    res["events"].sort(key=lambda e: e["date"])
    return res


def manual_entries(rows, base) -> list:
    """'기타매출' 시트의 금액을 상품군별 가상 품목으로 만든다 (수량=금액, 단가=1). 예측은 하지 않는다."""
    by = {}
    for r in rows:
        by.setdefault(r["group"], []).append(r)
    out = []
    for g, rs in by.items():
        rs = sorted(rs, key=lambda x: x["date"])
        past = [x["date"] for x in rs if x["date"] <= base]
        out.append(dict(
            name=f"{g} (직접입력)", item=None, price=0.0, Q=None, rate=0.0, r_po=0.0, r_sheet=None, rate_src="",
            gap=None, next=None, overdue=False, last=(past[-1] if past else None), fut_last=None, zero60=0,
            note="'기타매출' 시트 직접 입력 (예측 없음)", q_hist=[], consumed=None, start_stock=None, group=g, manual=True,
            recs=[dict(date=x["date"], name=g, qty=x["amount"], price=1.0, center="직접입력") for x in rs],
            events=[dict(date=x["date"], kind="실적" if x["date"] <= base else "확정",
                         demand=x["amount"], deliv=x["amount"], price=1.0) for x in rs]))
    return out


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


def won(v: float) -> str:
    if v is None:
        return "-"
    if abs(v) >= 1e8:
        return f"{v / 1e8:,.2f}억"
    return f"{v / 1e4:,.0f}만원"


# ─────────────────────────────── 테마 · 공통 UI ───────────────────────────────
C = dict(act="#1f4e79", conf="#2a9d8f", pred="#9ecae1", loss="#e76f51", rec="#f4a261", ly="#8a8f98",
         main="#1f4e79", vf="#f4a261", safe="#e9b949", inb="#2a9d8f", after="#2a9d8f", before="#b0b7c3")
GROUP_COLORS = ["#1f4e79", "#2a9d8f", "#e9b949", "#e76f51", "#8e6bbf", "#5aa9e6", "#9c6644", "#6c757d"]

CSS = """
<style>
.block-container {padding-top: 1.4rem; padding-bottom: 3rem; max-width: 1500px;}
[data-testid="stMetric"] {background: rgba(127,127,127,0.06); border: 1px solid rgba(127,127,127,0.18);
  border-radius: 12px; padding: 12px 16px;}
[data-testid="stMetricLabel"] p {font-size: 0.82rem; opacity: 0.78;}
[data-testid="stMetricValue"] {font-size: 1.4rem; font-weight: 700;}
[data-testid="stSidebar"] h3 {font-size: 1rem; margin-top: .6rem;}
.stTabs [data-baseweb="tab-list"] {gap: 4px;}
.stTabs [data-baseweb="tab"] {padding: 8px 14px; border-radius: 8px 8px 0 0; font-weight: 600;}
.kicker {font-size: .8rem; letter-spacing: .03em; opacity: .62; margin-bottom: .1rem;}
.sec {font-size: 1.05rem; font-weight: 700; margin: .4rem 0 .3rem 0;}
.muted {opacity: .65; font-size: .85rem;}
</style>
"""

TIER_BADGE = {0: "🔴 결품 불가피", 1: "🔴 오늘 발주", 1.5: "🟠 입고 전 결품", 2: "🟠 긴급", 3: "🟡 준비",
              4: "🟢 여유", 5: "⚪ 출고없음"}

toggle = getattr(st, "toggle", st.checkbox)


def badge(ev) -> str:
    if ev["tier"] == 4 and ev["d_pol"] is None:
        return "🟢 안정"
    return TIER_BADGE.get(ev["tier"], "")


def box():
    """테두리 있는 카드 컨테이너 (구버전 Streamlit은 일반 컨테이너)."""
    try:
        return st.container(border=True)
    except TypeError:
        return st.container()


def sec(title: str, sub: str = ""):
    st.markdown(f"<div class='sec'>{title}</div>" + (f"<div class='muted'>{sub}</div>" if sub else ""),
                unsafe_allow_html=True)


def style_fig(fig, height=None, legend=True, title=None):
    lay = dict(font=dict(family="Pretendard, 'Noto Sans KR', 'Malgun Gothic', sans-serif", size=12),
               margin=dict(l=10, r=10, t=44 if title else 16, b=10), hoverlabel=dict(font_size=12),
               showlegend=legend, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
    if legend:
        lay["legend"] = dict(orientation="h", y=-0.2, x=0)
    if title:
        lay["title"] = dict(text=title, font=dict(size=14))
    if height:
        lay["height"] = height
    fig.update_layout(**lay)
    fig.update_xaxes(gridcolor="rgba(128,128,128,0.15)")
    fig.update_yaxes(gridcolor="rgba(128,128,128,0.15)", zerolinecolor="rgba(128,128,128,0.3)")
    return fig


def show_progress(value: float, target: float, label: str):
    if not target:
        return
    ratio = value / target
    txt = f"{label} {won(value)} / 목표 {won(target)} · {ratio:.0%}"
    try:
        st.progress(min(max(ratio, 0.0), 1.0), text=txt)
    except TypeError:
        st.progress(min(max(ratio, 0.0), 1.0))
        st.caption(txt)


def pct_col(label, mx=150):
    try:
        return st.column_config.ProgressColumn(label, min_value=0, max_value=mx, format="%.0f%%")
    except Exception:
        return st.column_config.NumberColumn(label, format="%.0f%%")


def date_col(label):
    return st.column_config.DateColumn(label, format="MM/DD")


# ─────────────────────────────── 상품군 · 목표 ───────────────────────────────
DEFAULT_GROUP_RULES = [("기저귀", "반려견 기저귀"), ("장작", "장작"), ("필름", "보호필름"), ("필터", "차량용 필터"),
                       ("냄비", "주방(냄비·팬)"), ("팬", "주방(냄비·팬)"), ("웍", "주방(냄비·팬)")]


def group_of(name: str, rules) -> str:
    k = clean_txt(name)
    for kw, g in rules:
        if clean_txt(kw) and clean_txt(kw) in k:
            return g
    return "기타"


def total_target(targets: dict, month, groups=None) -> float:
    t = targets.get(month, {})
    if groups:
        return sum(t.get(g, 0.0) for g in groups)
    return t.get("전체") or sum(v for g, v in t.items() if g != "전체")


def months_span(fcs, base, horizon_end):
    ds = [e["date"] for f in fcs for e in f["events"]]
    first = min(ds) if ds else base
    out, (y, m) = [], (first.year, first.month)
    while (y, m) <= (horizon_end.year, horizon_end.month):
        out.append((y, m))
        y, m = add_months(y, m, 1)
    return out


def mkey(d):
    return d.year, d.month


def month_tot(fs, month, vat, kind=None, demand=False):
    return sum(ev_rev(e, vat, demand) for f in fs for e in f["events"]
               if mkey(e["date"]) == month and (kind is None or e["kind"] == kind))


def f_cat(f):
    return f["item"].cat if f["item"] else "미매칭"


def mlabel(t):
    return f"{t[0]}년 {t[1]}월"


# ─────────────────────────────── 매출 탭 ───────────────────────────────
def render_po_tab(ctx):
    fcs, base, horizon_end, vat_incl, cfg = ctx["fcs"], ctx["base"], ctx["horizon_end"], ctx["vat_incl"], ctx["po_cfg"]
    vat_txt = "VAT 포함" if vat_incl else "VAT 제외(공급가)"
    if not fcs:
        st.info("엑셀에 쿠팡 PO 이력 탭이 없습니다. 헤더에 '날짜'와 '단가'가 있는 탭을 넣으면 자동으로 인식합니다.")
        return
    months = ctx["months"]
    cur = (base.year, base.month)

    c1, c2, c3 = st.columns([1, 2, 2])
    with c1:
        sel = st.selectbox("📅 월", months, index=months.index(cur) if cur in months else len(months) - 1,
                           format_func=mlabel, key="rev_month")
    with c2:
        sel_g = st.multiselect("상품군 (비우면 전체)", sorted({f["group"] for f in fcs}), key="rev_groups")
    with c3:
        sel_c = st.multiselect("구분 (비우면 전체)", sorted({f_cat(f) for f in fcs}), key="rev_cats")
    view = [f for f in fcs if (not sel_g or f["group"] in sel_g) and (not sel_c or f_cat(f) in sel_c)]
    if not view:
        st.info("조건에 맞는 품목이 없습니다.")
        return
    sy, sm = sel
    ly = (sy - 1, sm)

    # ── KPI
    r_act, r_conf, r_pred = (month_tot(view, sel, vat_incl, k) for k in ("실적", "확정", "예측"))
    r_all = r_act + r_conf + r_pred
    r_dem = month_tot(view, sel, vat_incl, demand=True)
    r_ly = month_tot(view, ly, vat_incl, "실적")
    tgt = total_target(ctx["targets"], sel, sel_g or None)
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric(f"💰 {sm}월 예상 매출", won(r_all),
              delta=(f"전년 동월 대비 {(r_all / r_ly - 1) * 100:+.0f}%" if r_ly > 0 else None))
    k2.metric("✅ 확보 (실적+확정)", won(r_act + r_conf), help=f"실적 {won(r_act)} · 확정 {won(r_conf)}")
    k3.metric("🔮 예측 PO", won(r_pred))
    k4.metric("⛔ 재고부족 손실", won(r_dem - r_all), help="재고가 있었다면 더 받을 수 있었던 예측 PO 금액")
    k5.metric("🎯 목표 달성(예상)", f"{r_all / tgt:.0%}" if tgt else "목표 미입력",
              help="목표는 '데이터 설정' 탭 또는 엑셀 '목표' 시트에서 입력합니다.")
    if tgt:
        show_progress(r_act + r_conf, tgt, "확보")
        show_progress(r_all, tgt, "예상")
    st.caption(f"{vat_txt} · 기준일 {base:%Y-%m-%d} · 실적 = 기준일까지 PO, 확정 = 엑셀에 입력된 기준일 이후 PO, 예측 = 이력 기반 추정"
               + ("" if cfg["cap"] else " · 재고 제약 미반영(수요 기준)"))

    st.divider()
    render_group_section(view, sel, ctx, with_targets=not sel_g and not sel_c)

    # ── 품목별 표
    st.divider()
    sec(f"📋 품목별 {sm}월 예상 매출 · 차기 PO", "행을 선택하면 아래 'SKU 상세'가 그 품목으로 바뀝니다.")
    rows = []
    for f in view:
        it = f["item"]
        ev = [e for e in f["events"] if mkey(e["date"]) == sel]
        part = lambda k: [e for e in ev if e["kind"] == k]
        rev = lambda es, d=False: sum(ev_rev(e, vat_incl, d) for e in es)
        nxt = f["next"]
        rows.append({
            "상품군": f["group"], "품명(PO)": f["name"],
            "차기 PO": nxt, "D-": (nxt - base).days if nxt else None,
            "지연": "⏰" if f["overdue"] else "",
            "예상 매출": int(rev(ev)), "실적": int(rev(part("실적"))), "확정": int(rev(part("확정"))),
            "예측": int(rev(part("예측"))), "재고부족 손실": int(rev(ev, True) - rev(ev)),
            "전년 동월": int(sum(ev_rev(e, vat_incl) for e in f["events"] if e["kind"] == "실적" and mkey(e["date"]) == ly)),
            "구분": f_cat(f), "유형": ("밀크런" if it.is_milkrun else "VF") if it else "-",
            "단가": int(f["price"]), "PO 1회 수량": int(f["Q"]) if f["Q"] else None,
            "일 출고율": round(f["rate"], 1) if f["rate"] else None,
            "PO 간격(일)": round(f["gap"], 1) if f["gap"] else None,
            "출고율 근거": f["rate_src"] or f["note"],
            "마지막 PO": f["last"], "확정 PO ~": f["fut_last"],
            "PO 건수": len(ev), "수량": int(sum(e["deliv"] for e in ev)),
            "0납품 PO(60일)": f["zero60"] or None,
        })
    df = pd.DataFrame(rows).sort_values(["예상 매출", "D-"], ascending=[False, True]).reset_index(drop=True)
    show_all = toggle("예측 근거 컬럼까지 보기", False, key="po_all_cols")
    core = ["상품군", "품명(PO)", "차기 PO", "D-", "지연", "예상 매출", "실적", "확정", "예측", "재고부족 손실", "전년 동월"]
    show = df if show_all else df[core]
    money = ["예상 매출", "실적", "확정", "예측", "재고부족 손실", "전년 동월", "단가"]
    ccfg = {c: num_col(c) for c in money + ["PO 1회 수량", "수량"]}
    ccfg.update({c: date_col(c) for c in ("차기 PO", "마지막 PO", "확정 PO ~")})
    names = list(df["품명(PO)"])
    if st.session_state.get("po_pick") not in names:
        st.session_state["po_pick"] = names[0]
    if _HAS_SELECT:
        event = wide(st.dataframe, show, hide_index=True, column_config=ccfg,
                     on_select="rerun", selection_mode="single-row", key="po_tbl")
        rows_sel = list(event.selection.rows) if event is not None and event.selection is not None else []
        picked = str(show.iloc[rows_sel[0]]["품명(PO)"]) if rows_sel else None
        if picked and picked != st.session_state.get("_last_po_row"):
            st.session_state["_last_po_row"] = picked
            st.session_state["po_pick"] = picked
        elif not picked:
            st.session_state["_last_po_row"] = None
    else:
        wide(st.dataframe, show, hide_index=True, column_config=ccfg)
    st.caption("차기 PO: VF형은 마지막 PO 이후 실제 출고 누적이 'PO 1회 수량'을 채우는 날, 밀크런은 마지막(확정 포함) PO + PO 간격. "
               "⏰ = 예상일이 지났는데 PO가 없음 (재고 부족·쿠팡 발주 보류 등 확인 필요)."
               + (" 토·일 예측분은 월요일로 옮깁니다." if cfg["weekend"] else ""))

    render_loss(view, base, sel, horizon_end, vat_incl, cfg, ctx["deadlines"])

    # ── 월별 추이
    st.divider()
    sec("📈 월별 매출 추이")
    mode = st.radio("구성", ["실적·확정·예측", "상품군별"], horizontal=True, key="trend_mode")
    y0, m0 = add_months(cur[0], cur[1], -12)
    cm = [mm for mm in months if mm >= (y0, m0)]
    labels = [f"{a % 100:02d}.{b:02d}" for a, b in cm]
    fig = go.Figure()
    if mode.startswith("실적"):
        for k in ("실적", "확정", "예측"):
            fig.add_trace(go.Bar(x=labels, y=[month_tot(view, mm, vat_incl, k) for mm in cm],
                                 name=k if k != "예측" or not cfg["cap"] else "예측(재고 반영)",
                                 marker_color=C[{"실적": "act", "확정": "conf", "예측": "pred"}[k]],
                                 hovertemplate="%{y:,.0f}원"))
        if cfg["cap"]:
            fig.add_trace(go.Scatter(x=labels, y=[month_tot(view, mm, vat_incl, demand=True) for mm in cm],
                                     name="수요 기준(재고 무제한)", mode="lines+markers",
                                     line=dict(color=C["loss"], dash="dash"), hovertemplate="%{y:,.0f}원"))
    else:
        for i, g in enumerate(sorted({f["group"] for f in view})):
            fs = [f for f in view if f["group"] == g]
            fig.add_trace(go.Bar(x=labels, y=[month_tot(fs, mm, vat_incl) for mm in cm], name=g,
                                 marker_color=GROUP_COLORS[i % len(GROUP_COLORS)], hovertemplate="%{y:,.0f}원"))
    fig.add_trace(go.Scatter(x=labels, y=[month_tot(view, (mm[0] - 1, mm[1]), vat_incl, "실적") for mm in cm],
                             name="전년 동월 실적", mode="lines+markers", line=dict(color=C["ly"], dash="dot"),
                             hovertemplate="%{y:,.0f}원"))
    tg = [total_target(ctx["targets"], mm, sel_g or None) for mm in cm]
    if any(tg):
        fig.add_trace(go.Scatter(x=labels, y=[t or None for t in tg], name="목표", mode="markers",
                                 marker=dict(symbol="line-ew-open", size=26, color="#222", line=dict(width=3)),
                                 hovertemplate="목표 %{y:,.0f}원"))
    fig.update_layout(barmode="stack", hovermode="x unified", yaxis_title=f"매출 (원, {vat_txt})")
    wide(st.plotly_chart, style_fig(fig, height=380))

    # ── SKU 상세
    st.divider()
    sec("🔍 SKU 상세 · 예상 매출 · PO 일정 · 예측 근거")
    pick = st.selectbox("품목 선택:", names, key="po_pick")
    with box():
        render_po_detail(next(x for x in view if x["name"] == pick), base, sel, months, vat_incl, cfg)


def group_rows(view, sel, ctx, with_targets=True):
    vat = ctx["vat_incl"]
    prev, ly = add_months(sel[0], sel[1], -1), (sel[0] - 1, sel[1])
    rows = []
    gs = {f["group"] for f in view}
    if with_targets:      # 매출 데이터는 없어도 목표가 있는 상품군(예: 네이버 두부모래)도 표에 보이게
        gs |= {g for g, v in ctx["targets"].get(sel, {}).items() if g != "전체" and v}
    for g in sorted(gs):
        fs = [f for f in view if f["group"] == g]
        a, c, p = (month_tot(fs, sel, vat, k) for k in ("실적", "확정", "예측"))
        tot = a + c + p
        dem = month_tot(fs, sel, vat, demand=True)
        t = ctx["targets"].get(sel, {}).get(g, 0.0)
        pv, lyv = month_tot(fs, prev, vat), month_tot(fs, ly, vat, "실적")
        rows.append({"상품군": g, "예상 매출": int(tot), "실적": int(a), "확정": int(c), "예측": int(p),
                     "목표": int(t) if t else None,
                     "달성률(예상)": round(tot / t * 100) if t else None, "확보율": round((a + c) / t * 100) if t else None,
                     "전월": int(pv), "전월비(%)": round((tot / pv - 1) * 100) if pv else None,
                     "전년 동월": int(lyv), "전년비(%)": round((tot / lyv - 1) * 100) if lyv else None,
                     "재고부족 손실": int(dem - tot), "SKU": len(fs)})
    df = pd.DataFrame(rows)
    df["데이터"] = ["" if n else "매출 데이터 없음 ('기타매출' 시트에 입력)" for n in df["SKU"]]
    return df.sort_values(["예상 매출", "목표"], ascending=False, na_position="last").reset_index(drop=True)


def won_s(v) -> str:
    """표 안에 넣을 짧은 금액 문자열 (값이 없으면 '-')."""
    if v is None or (isinstance(v, float) and math.isnan(v)) or abs(v) < 0.5:
        return "-"
    if abs(v) >= 1e8:
        return f"{v / 1e8:,.2f}억"
    return f"{v / 1e4:,.0f}만"


def pct_s(v, sign=False) -> str:
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return "-"
    return f"{v:+.0f}%" if sign else f"{v:.0f}%"


def render_group_section(view, sel, ctx, with_targets=True):
    sec(f"📦 상품군별 {sel[1]}월 매출", "확보 = 실적 + 확정 · 상품군 기준은 '데이터 설정' 탭에서 바꿀 수 있습니다.")
    df = group_rows(view, sel, ctx, with_targets)
    nz = lambda v: None if v is None or (isinstance(v, float) and math.isnan(v)) else v
    show = pd.DataFrame({
        "상품군": df["상품군"],
        "예상 매출": df["예상 매출"].map(won_s),
        "목표": df["목표"].map(lambda v: won_s(nz(v))),
        "달성률": df["달성률(예상)"],
        "확보 매출": (df["실적"] + df["확정"]).map(won_s),
        "확보율": df["확보율"],
        "실적": df["실적"].map(won_s), "확정": df["확정"].map(won_s), "예측": df["예측"].map(won_s),
        "전월": df["전월"].map(won_s), "전월비": df["전월비(%)"].map(lambda v: pct_s(nz(v), True)),
        "전년 동월": df["전년 동월"].map(won_s), "전년비": df["전년비(%)"].map(lambda v: pct_s(nz(v), True)),
        "재고부족 손실": df["재고부족 손실"].map(lambda v: won_s(v) if v else "-"),
        "SKU": df["SKU"], "비고": df["데이터"],
    })
    ccfg = {"달성률": pct_col("달성률"), "확보율": pct_col("확보율")}
    try:
        ccfg["상품군"] = st.column_config.TextColumn("상품군", pinned=True)
    except TypeError:
        pass
    wide(st.dataframe, show, hide_index=True, column_config=ccfg)

    tot = df["실적"] + df["확정"] + df["예측"]
    tgt = df["목표"].map(nz)
    fig = go.Figure()
    y = list(df["상품군"])[::-1]
    for k in ("실적", "확정", "예측"):
        fig.add_trace(go.Bar(y=y, x=list(df[k])[::-1], orientation="h", name=k,
                             marker_color=C[{"실적": "act", "확정": "conf", "예측": "pred"}[k]],
                             hovertemplate=k + " %{x:,.0f}원<extra></extra>"))
    if tgt.notna().any():
        fig.add_trace(go.Scatter(y=y, x=list(tgt)[::-1], mode="markers", name="목표",
                                 marker=dict(symbol="line-ns-open", size=22, color="#222", line=dict(width=3)),
                                 hovertemplate="목표 %{x:,.0f}원<extra></extra>"))
    labels = [won_s(t) + (f" / 목표 {won_s(g)} ({t / g:.0%})" if g else "") for t, g in zip(tot, tgt)]
    xs = [max(t, g or 0) for t, g in zip(tot, tgt)]
    fig.add_trace(go.Scatter(y=y, x=xs[::-1], mode="text", text=labels[::-1], textposition="middle right",
                             textfont=dict(size=12), showlegend=False, hoverinfo="skip"))
    fig.update_layout(barmode="stack", xaxis=dict(range=[0, max(xs + [1]) * 1.45]))
    wide(st.plotly_chart, style_fig(fig, height=max(220, 44 * len(df) + 90)))


def render_loss(view, base, sel, horizon_end, vat_incl, cfg, deadlines):
    """재고 부족으로 납품 못 할 PO(확정·예측)를 품목별로 집계: 얼마나, 언제부터, 신규 발주로 막을 수 있는지."""
    st.divider()
    st.subheader("⛔ 재고 부족으로 빠지는 매출 (품목별)")
    if not cfg["cap"]:
        st.info("사이드바에서 '재고 제약 반영'을 켜면 재고 부족으로 빠지는 매출을 볼 수 있습니다.")
        return
    sy, sm = sel
    scope = st.radio("집계 범위", [f"{sy}년 {sm}월", f"예측 기간 전체 (~{horizon_end:%Y-%m})"], horizontal=True, key="loss_scope")
    in_scope = (lambda d: (d.year, d.month) == sel) if scope.startswith(f"{sy}년") else (lambda d: d > base)
    k = 1.0 if vat_incl else 1 / 1.1
    rows = []
    for f in view:
        it = f["item"]
        sh = [e for e in f["events"] if e["kind"] != "실적" and in_scope(e["date"]) and e["demand"] - e["deliv"] > 1e-9]
        rk = [e for e in f["events"] if e["kind"] == "확정" and in_scope(e["date"]) and e.get("risk", 0) > 1e-9]
        if not sh and not rk:
            continue
        lost = lambda es: sum((e["demand"] - e["deliv"]) * e["price"] for e in es) * k
        first = min(e["date"] for e in sh + rk)
        lt = it.lt if it else None
        cut = base + timedelta(days=lt) if lt else None          # 오늘 공장 발주해도 이 날 전에는 입고 불가
        unrec = [e for e in sh if cut is None or e["date"] < cut]
        rec = [e for e in sh if cut is not None and e["date"] >= cut]
        ev = deadlines.get(it.name) if it else None
        inb = ", ".join(f"{d:%m/%d} {q:,.0f}" for d, q, _ in it.inbound) if it and it.inbound else "-"
        rows.append({
            "구분": it.cat if it else "미매칭", "품명(PO)": f["name"],
            "빠지는 매출": int(lost(sh)), "빠지는 수량": int(sum(e["demand"] - e["deliv"] for e in sh)),
            "미납 PO 건수": len(sh), "첫 미납일": first.strftime("%m/%d" if first.year == base.year else "%y/%m/%d"),
            "평택 재고": int(it.main) if it else None, "발주완료 입고": inb,
            "└ 리드타임 내(못 막음)": int(lost(unrec)), "└ 신규 발주로 회복 가능": int(lost(rec)),
            "확정 PO 납품 위험": int(sum(e["risk"] * e["price"] for e in rk) * k),
            "리드타임": lt,
            "공장발주 데드라인": (ev["deadline"].strftime("%m/%d") if ev and ev["deadline"] else "-"),
        })
    if not rows:
        st.success("이 범위에서는 재고 부족으로 빠지는 매출이 없습니다.")
        return
    df = pd.DataFrame(rows)
    df = df.assign(_s=df["빠지는 매출"] + df["확정 PO 납품 위험"]).sort_values("_s", ascending=False).drop(columns="_s").reset_index(drop=True)
    t, u, r = df["빠지는 매출"].sum(), df["└ 리드타임 내(못 막음)"].sum(), df["└ 신규 발주로 회복 가능"].sum()
    rr = df["확정 PO 납품 위험"].sum()
    a, b, c, d4 = st.columns(4)
    a.metric("빠지는 매출 합계", won(t), help=f"{len(df)}개 품목")
    b.metric("리드타임 안이라 못 막는 금액", won(u), help="오늘 공장 발주해도 리드타임 전에는 입고가 안 되는 기간의 손실")
    c.metric("지금 발주하면 회복 가능한 금액", won(r))
    d4.metric("확정 PO 중 재고 없는 금액", won(rr),
              help="이미 받은 확정 PO인데 평택 재고·발주완료 입고분으로는 수량이 안 나오는 금액. 예상 매출에는 그대로 포함돼 있으니 납품 가능 여부를 확인하세요.")

    fig = go.Figure()
    lab = [n if len(n) <= 24 else n[:23] + "…" for n in df["품명(PO)"]][::-1]
    fig.add_trace(go.Bar(y=lab, x=list(df["└ 리드타임 내(못 막음)"])[::-1], orientation="h", name="리드타임 내(못 막음)",
                         marker_color="#e76f51", hovertemplate="%{x:,.0f}원"))
    fig.add_trace(go.Bar(y=lab, x=list(df["└ 신규 발주로 회복 가능"])[::-1], orientation="h", name="신규 발주로 회복 가능",
                         marker_color="#f4a261", hovertemplate="%{x:,.0f}원"))
    fig.update_layout(barmode="stack", height=max(260, 34 * len(df) + 80), margin=dict(l=20, r=20, t=20, b=20),
                      xaxis_title=f"원 ({'VAT 포함' if vat_incl else 'VAT 제외'})", legend=dict(orientation="h", y=-0.15))
    wide(st.plotly_chart, style_fig(fig, height=fig.layout.height))

    ccfg = {c_: num_col(c_) for c_ in ("빠지는 매출", "빠지는 수량", "평택 재고", "└ 리드타임 내(못 막음)", "└ 신규 발주로 회복 가능", "확정 PO 납품 위험")}
    if _HAS_SELECT:
        event = wide(st.dataframe, df, hide_index=True, column_config=ccfg,
                     on_select="rerun", selection_mode="single-row", key="loss_tbl")
        rs = list(event.selection.rows) if event is not None and event.selection is not None else []
        picked = str(df.iloc[rs[0]]["품명(PO)"]) if rs else None
        if picked and picked != st.session_state.get("_last_loss_row"):
            st.session_state["_last_loss_row"] = picked
            st.session_state["po_pick"] = picked          # 아래 SKU 상세를 이 품목으로
        elif not picked:
            st.session_state["_last_loss_row"] = None
    else:
        wide(st.dataframe, df, hide_index=True, column_config=ccfg)
    st.caption("빠지는 매출 = 확정·예측 PO 중 평택 재고(+발주완료 입고분)가 모자라 납품 못 할 수량 × 단가. "
               "첫 미납일 = 재고가 처음 모자라는 PO 날짜. 리드타임 내 = 오늘 공장 발주해도 '기준일+리드타임' 전이라 막을 수 없는 손실, "
               "그 이후 손실은 지금 발주하면 회복 가능. 확정 PO 납품 위험 = 이미 받은 PO라 매출에는 넣었지만 재고상 수량이 안 나오는 금액. 공장발주 데드라인은 '재고·공장 발주' 탭 기준입니다. "
               "행을 선택하면 아래 SKU 상세에서 PO별 미납 수량을 볼 수 있습니다.")


def render_po_detail(f, base, sel, months, vat_incl, cfg):
    if f.get("manual"):
        ev = [e for e in f["events"] if (e["date"].year, e["date"].month) == sel]
        st.markdown(f"**{f['name']}** · 엑셀 '기타매출' 시트에 직접 입력한 매출 (쿠팡 PO 아님, 예측 없음)")
        st.metric(f"{sel[1]}월 매출", won(sum(ev_rev(e, vat_incl) for e in ev)))
        wide(st.dataframe, pd.DataFrame([{"일자": x["date"], "매출": int(ev_rev(x, vat_incl))} for x in f["events"]][::-1]),
             hide_index=True, column_config={"일자": date_col("일자"), "매출": num_col("매출")})
        return
    it = f["item"]
    sy, sm = sel
    mk = lambda d: (d.year, d.month)
    rev = lambda es, d=False: sum(ev_rev(e, vat_incl, d) for e in es)
    ev = [e for e in f["events"] if mk(e["date"]) == sel]
    part = lambda k: [e for e in ev if e["kind"] == k]
    ly_rev = rev([e for e in f["events"] if e["kind"] == "실적" and mk(e["date"]) == (sy - 1, sm)])
    total = rev(ev)
    nxt = f["next"]

    st.markdown(f"**{f['name']}** · "
                + (f"{it.cat} / {'밀크런' if it.is_milkrun else 'VF형'} · 재고시트: {it.name}" if it else "재고시트 매칭 없음")
                + f" · 단가 {f['price']:,.0f}원")
    a, b, c, d, e5 = st.columns(5)
    a.metric(f"{sm}월 예상 매출", won(total),
             delta=(f"전년 동월 대비 {(total / ly_rev - 1) * 100:+.0f}%" if ly_rev > 0 else None))
    b.metric("실적 / 확정", f"{won(rev(part('실적')))} / {won(rev(part('확정')))}")
    c.metric("예측", won(rev(part("예측"))), help=f"예측 PO {len(part('예측'))}건")
    d.metric("재고 부족 차감", won(rev(ev, True) - total))
    e5.metric("차기 PO 예상", (nxt.strftime("%m/%d") + (" ⏰" if f["overdue"] else "")) if nxt else "-",
              delta=(f"D-{(nxt - base).days}" if nxt else None), delta_color="off")

    # 예측 근거
    lines = []
    if f["Q"]:
        lines.append(f"PO 1회 수량 **{f['Q']:,.0f}개** = 최근 {len(f['q_hist'])}회 PO 수량 "
                     f"({', '.join(f'{q:,.0f}' for q in f['q_hist'])})의 중앙값")
    if f["rate"]:
        lines.append(f"일 출고율 **{f['rate']:,.1f}개/일** ({f['rate_src']}"
                     + (f", 수요 보정 {cfg['mult'] * 100:.0f}% 적용" if cfg["mult"] != 1 else "") + ")"
                     + f" · 참고: PO 실적 {f['r_po']:,.1f}개/일"
                     + (f", 시트 출고 {f['r_sheet']:,.1f}개/일" if f["r_sheet"] is not None else ""))
    if f["gap"]:
        lines.append(f"PO 간격 **{f['gap']:,.1f}일** = {f['Q']:,.0f} ÷ {f['rate']:,.1f}")
    if f["fut_last"]:
        lines.append(f"엑셀에 이미 잡힌 확정 PO가 {f['fut_last']:%m/%d}까지 있어 그 다음부터 예측")
    elif f["consumed"] is not None and f["last"]:
        lines.append(f"마지막 PO {f['last']:%m/%d} 이후 시트 출고 누적 **{f['consumed']:,.0f}개** / PO 1회 {f['Q']:,.0f}개"
                     + (" → 이미 채워졌는데 PO가 없어 ⏰지연" if f["overdue"] else " → 남은 만큼 출고되는 날을 차기 PO로"))
    elif f["last"] and f["gap"]:
        lines.append(f"마지막 PO {f['last']:%m/%d} + 간격 {f['gap']:,.1f}일")
    if it is not None and cfg["cap"]:
        inb = [(dd, q) for dd, q, _ in it.inbound]
        lines.append(f"납품 가능 재고: 평택 {it.main:,.0f}개"
                     + (" + 발주완료 입고 " + ", ".join(f"{dd:%m/%d} {q:,.0f}개" for dd, q in inb) if inb else " (발주완료 입고 예정 없음)"))
    elif it is None:
        lines.append("재고시트와 매칭되지 않아 재고 제약 없이 PO 실적만으로 계산")
    if f["note"]:
        lines.append(f"⚠️ {f['note']}")
    if f["zero60"]:
        lines.append(f"최근 60일 0개 납품 PO {f['zero60']}건 (재고 부족 등으로 못 받은 PO)")
    st.markdown("\n".join(f"- {x}" for x in lines))

    g1, g2 = st.columns(2)
    with g1:   # PO 타임라인 (최근 90일 ~ 예측 범위)
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
        fig2.update_layout(title="PO 타임라인 (수량)", barmode="stack", margin=dict(l=20, r=20, t=40, b=20),
                           legend=dict(orientation="h", y=-0.2))
        wide(st.plotly_chart, style_fig(fig2, height=320, title="PO 타임라인 (수량)"))
    with g2:   # 이 SKU의 월별 매출
        y0, m0 = add_months(base.year, base.month, -12)
        cm = [mm for mm in months if mm >= (y0, m0)]
        labels = [f"{a_ % 100:02d}.{b_:02d}" for a_, b_ in cm]
        fig3 = go.Figure()
        for k, col in (("실적", "#1f4e79"), ("확정", "#2a9d8f"), ("예측", "#9ecae1")):
            fig3.add_trace(go.Bar(x=labels, marker_color=col, name=k, hovertemplate="%{y:,.0f}원",
                                  y=[rev([e for e in f["events"] if e["kind"] == k and mk(e["date"]) == mm]) for mm in cm]))
        fig3.add_trace(go.Scatter(x=labels, name="전년 동월", mode="lines+markers", line=dict(color="#888", dash="dot"),
                                  hovertemplate="%{y:,.0f}원",
                                  y=[rev([e for e in f["events"] if e["kind"] == "실적" and mk(e["date"]) == (mm[0] - 1, mm[1])]) for mm in cm]))
        fig3.update_layout(title="월별 매출", barmode="stack", hovermode="x unified", margin=dict(l=20, r=20, t=40, b=20),
                           legend=dict(orientation="h", y=-0.2))
        wide(st.plotly_chart, style_fig(fig3, height=320, title="월별 매출"))

    # 선택월 PO 일정
    st.markdown(f"**{sy}년 {sm}월 PO 일정**")
    if ev:
        wd = "월화수목금토일"
        tbl = pd.DataFrame([{
            "일자": e["date"].strftime("%m/%d") + f"({wd[e['date'].weekday()]})", "구분": e["kind"],
            "PO 수량": int(e["demand"]), "납품(예상)": int(e["deliv"]), "단가": int(e["price"]),
            "매출": int(ev_rev(e, vat_incl)),
            "평택 잔여": int(e["stock_after"]) if "stock_after" in e else None,
            "비고": (f"재고 부족 {e['demand'] - e['deliv']:,.0f}개 미납" if e["demand"] - e["deliv"] > 1e-9
                     else f"⚠️ 재고 {e['risk']:,.0f}개 모자람 (확정 PO, 매출엔 포함)" if e.get("risk", 0) > 1e-9 else ""),
        } for e in ev])
        wide(st.dataframe, tbl, hide_index=True,
             column_config={k: num_col(k) for k in ("PO 수량", "납품(예상)", "단가", "매출", "평택 잔여")})
    else:
        st.write("이 달에는 실적·확정·예측 PO가 없습니다.")

    with st.expander("📜 PO 원본 이력 (최근 60건, 0개 납품 포함)"):
        wd = "월화수목금토일"
        hist = [x for x in f["recs"]][-60:][::-1]
        wide(st.dataframe, pd.DataFrame([{
            "일자": x["date"].strftime("%Y-%m-%d") + f"({wd[x['date'].weekday()]})", "수량": int(x["qty"]),
            "단가": int(x["price"]), "매출": int(x["qty"] * x["price"] * (1 if vat_incl else 1 / 1.1)),
            "센터": x["center"], "상태": "확정(미래)" if x["date"] > base else ("0개 납품" if x["qty"] <= 0 else "")}
            for x in hist]), hide_index=True,
             column_config={k: num_col(k) for k in ("수량", "단가", "매출")})


# ─────────────────────────────── 재고 · 공장 발주 탭 ───────────────────────────────
def inventory_table(ctx) -> pd.DataFrame:
    base, rows = ctx["base"], []
    for it in ctx["items"]:
        sim, ev = ctx["calc"][it.name]
        total = it.main + it.vf
        d = ev["deadline"]
        off = lambda i: (base + timedelta(days=i)) if i is not None else None
        rows.append({
            "_tier": ev["tier"], "_dpol": ev["d_pol"] if ev["d_pol"] is not None else 9999,
            "상태": badge(ev), "품명": it.name, "상품군": ctx["item_group"][it.name], "구분": it.cat,
            "유형": "밀크런" if it.is_milkrun else "VF",
            "평택재고": int(it.main), "VF재고": None if it.is_milkrun else int(it.vf),
            "안전재고": int(it.safety) if it.safety > 0 else None,
            "안전재고 대비": round(it.main / it.safety * 100) if it.safety > 0 else None,
            "일평균출고": round(it.adu, 1), "재고커버(일)": round(total / it.adu) if it.adu > 0 else None,
            "발주 데드라인": d, "D-": (d - base).days if d else None,
            "권장 발주량": ev["rec_qty"] if it.adu > 0 and ev["rec_qty"] else None,
            "평택 결품일": off(sim["stockout"]), "안전재고 이탈일": off(ev["safety_day"]),
            "입고예정(발주완료)": int(sum(q for _, q, _ in it.inbound)) or None,
            "입고계획(미발주)": int(sum(x["qty"] for x in it.inbound_detail if not x["ordered"])) or None,
            "차기 PO(이력예측)": ctx["next_po"].get(it.name),
            "차기 쿠팡PO(시뮬)": sim["first_po"],
            "출고예정(합)": int(sum(it.plan.values())) or None,
            "적용L/T(일)": it.lt, "L/T 근거": it.lt_src,
            "상세": ev["status"],
        })
    return pd.DataFrame(rows).sort_values(["_tier", "_dpol"]).reset_index(drop=True)


INV_CORE = ["상태", "품명", "상품군", "평택재고", "안전재고 대비", "재고커버(일)", "발주 데드라인", "D-", "권장 발주량",
            "평택 결품일", "입고예정(발주완료)", "차기 PO(이력예측)", "상세"]


def inv_colcfg():
    ccfg = {c: num_col(c) for c in ("평택재고", "VF재고", "안전재고", "권장 발주량", "입고예정(발주완료)",
                                     "입고계획(미발주)", "출고예정(합)")}
    ccfg.update({c: date_col(c) for c in ("발주 데드라인", "평택 결품일", "안전재고 이탈일", "차기 PO(이력예측)", "차기 쿠팡PO(시뮬)")})
    ccfg["안전재고 대비"] = pct_col("안전재고 대비", 200)
    ccfg["상세"] = st.column_config.TextColumn("상세", width="large")
    return ccfg


def render_inventory(ctx):
    base, items, df = ctx["base"], ctx["items"], ctx["inv_df"]
    c1, c2, c3 = st.columns([2, 2, 1])
    with c1:
        g_sel = st.multiselect("상품군 (비우면 전체)", sorted(df["상품군"].unique()), key="inv_groups")
    with c2:
        c_sel = st.multiselect("구분 (비우면 전체)", sorted(df["구분"].unique()), key="inv_cats")
    with c3:
        only = toggle("조치 필요만", False, key="inv_only", help="오늘 발주·긴급·준비·입고 전 결품 품목만 표시")
    view = df
    if g_sel:
        view = view[view["상품군"].isin(g_sel)]
    if c_sel:
        view = view[view["구분"].isin(c_sel)]
    if only:
        view = view[view["_tier"] <= 3]
    view = view.reset_index(drop=True)

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("🔴 오늘 발주 / 결품 불가피", f"{int((view['_tier'] <= 1).sum())}개")
    k2.metric("🟠 입고 전 결품", f"{int((view['_tier'] == 1.5).sum())}개", help="발주완료 입고분이 오기 전 잠깐 비는 품목 (입고·출고 일정 점검)")
    k3.metric("🟠🟡 14일 내 발주", f"{int(view['_tier'].isin([2, 3]).sum())}개")
    k4.metric("🟢 여유·안정", f"{int((view['_tier'] == 4).sum())}개")

    sec("📋 공장 발주 데드라인", "행을 선택하면 아래 '개별 SKU'가 그 품목으로 바뀝니다.")
    show_all = toggle("모든 컬럼 보기", False, key="inv_all_cols")
    show = view.drop(columns=["_tier", "_dpol"]) if show_all else view[INV_CORE]
    names = list(view["품명"])
    if not names:
        st.info("조건에 맞는 품목이 없습니다.")
        return
    if _HAS_SELECT:
        event = wide(st.dataframe, show, hide_index=True, column_config=inv_colcfg(),
                     on_select="rerun", selection_mode="single-row", key="master_tbl")
        rows_sel = list(event.selection.rows) if event is not None and event.selection is not None else []
        picked = str(show.iloc[rows_sel[0]]["품명"]) if rows_sel else None
        if picked and picked != st.session_state.get("_last_row_pick"):
            st.session_state["_last_row_pick"] = picked
            st.session_state[SKU_KEY] = picked
        elif not picked:
            st.session_state["_last_row_pick"] = None
    else:
        wide(st.dataframe, show, hide_index=True, column_config=inv_colcfg())
    with st.expander("ℹ️ 상태·데드라인 계산 기준"):
        st.markdown(
            "- **발주 데드라인** = (발주완료 입고분을 반영하고도 끝까지 회복되지 않는 안전재고 이탈일/결품일) − 리드타임\n"
            "- **🟠 입고 전 결품**: 발주완료 입고분으로 곧 회복되지만 그 전에 잠깐 비는 경우. 신규 발주가 아니라 입고 앞당기기·출고 조정 대상\n"
            "- **권장 발주량**: 신규분이 도착한 뒤 산정 기간 동안 안전재고(미기재 품목은 결품 방지) 기준 최대 부족분을 납품 MOQ 배수로 올림\n"
            "- **차기 PO(이력예측)**은 쿠팡PO 탭 이력 기반, **차기 쿠팡PO(시뮬)**은 사이드바 트리거 가정 기반")

    # ── 개별 SKU
    st.divider()
    sec("🔍 개별 SKU · 쿠팡 PO 차감 타임라인 & 재고 곡선")
    if st.session_state.get(SKU_KEY) not in names:
        st.session_state[SKU_KEY] = names[0]
    sku = st.selectbox("품목 선택:", names, key=SKU_KEY)
    it = next(i for i in items if i.name == sku)
    h = hashlib.md5(sku.encode()).hexdigest()[:6]
    trig_mode, trig_value = ctx["trig_mode"], ctx["trig_value"]
    with st.expander("이 품목만 가정 바꿔보기", expanded=False):
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        with c1:
            if trig_mode == "days":
                p_trig = st.number_input("PO 트리거 (VF 잔여일)", 1, 60, int(trig_value), key=f"tg_{h}")
            else:
                p_trig = st.number_input("PO 트리거 (VF 잔여 수량)", 0, 100000, int(trig_value), step=10, key=f"tg_{h}")
        with c2:
            p_mult = st.selectbox("쿠팡 PO 배수", [1, 2, 3], index=[1, 2, 3].index(ctx["po_mult"]), key=f"mu_{h}",
                                  format_func=lambda x: "전량" if it.moq_all else f"{x}배수 ({it.moq * x:,.0f}개)")
        with c3:
            p_adu = st.number_input("일일 출고량", 0.0, 1e6, float(round(it.adu, 2)), step=1.0, key=f"adu_{h}")
        with c4:
            p_safe = st.number_input("안전재고 (0=결품 기준)", 0, 10_000_000, int(it.safety), step=100, key=f"sf_{h}")
        with c5:
            p_lt = st.number_input("리드타임 (일)", 1, 120, int(it.lt), key=f"lt_{h}")
        with c6:
            p_tr = st.number_input("평택→VF 이동일", 0, 14, int(ctx["transfer_days"]), key=f"tr_{h}") if not it.is_milkrun else 0
        st.caption(f"L/T 근거: {it.lt_src} · 출고평균 {it.adu_days}일 기준"
                   + (f" · {it.plan_until:%m/%d}까지는 시트 입력 예정 출고 사용" if it.plan_until else ""))

    sim = simulate(it, base, p_adu, trig_mode, p_trig, p_mult, p_tr, log=True)
    ev = evaluate(it, sim, base, p_adu, p_safe, p_lt, ctx["target_days"])
    with box():
        st.markdown(f"**{badge(ev)}** · {ev['status']}")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("평택 재고", f"{it.main:,.0f}개", help=f"VF {it.vf:,.0f}개" if not it.is_milkrun else None)
        m2.metric("공장발주 데드라인", ev["deadline"].strftime("%Y-%m-%d") if ev["deadline"] else "-",
                  delta=(f"D-{(ev['deadline'] - base).days}" if ev["deadline"] else None), delta_color="off")
        m3.metric("평택 결품 예상일", fmt_md(base, sim["stockout"]))
        m4.metric("권장 발주량", f"{ev['rec_qty']:,}개" if p_adu > 0 else "-")

        days = [base + timedelta(days=i) for i in range(CHART_DAYS)]
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=days, y=sim["main"][:CHART_DAYS], mode="lines", name="평택 본창고",
                                 line=dict(color=C["main"], width=3), fill="tozeroy", fillcolor="rgba(31,78,121,0.08)"))
        if not it.is_milkrun:
            fig.add_trace(go.Scatter(x=days, y=sim["vf"][:CHART_DAYS], mode="lines", name="VF 잔여",
                                     line=dict(color=C["vf"], width=2, dash="dash")))
        arr = [(d, q) for d, q, _ in it.inbound if base <= d < days[-1]]
        if arr:
            fig.add_trace(go.Scatter(x=[d for d, _ in arr], y=[sim["main"][(d - base).days] for d, _ in arr],
                                     mode="markers", name="공장 입고 (발주완료)", text=[f"+{q:,.0f}개" for _, q in arr],
                                     hovertemplate="%{x|%m/%d} 입고 %{text}",
                                     marker=dict(symbol="triangle-up", size=13, color=C["inb"])))
        plan_arr = [x for x in it.inbound_detail if not x["ordered"] and base <= x["arrival"] < days[-1]]
        if plan_arr:
            fig.add_trace(go.Scatter(x=[x["arrival"] for x in plan_arr],
                                     y=[sim["main"][(x["arrival"] - base).days] for x in plan_arr],
                                     mode="markers", name="입고 계획 (미발주·미반영)",
                                     marker=dict(symbol="triangle-up-open", size=13, color="#888")))
        if p_safe > 0:
            fig.add_hline(y=p_safe, line_dash="dashdot", line_color=C["safe"], annotation_text=f"안전재고 {p_safe:,}")
        if ev["deadline"] and base <= ev["deadline"] < days[-1]:
            fig.add_vline(x=ev["deadline"].isoformat(), line_dash="dot", line_color=C["loss"])
        fig.update_layout(hovermode="x unified", yaxis_title="수량 (개)")
        wide(st.plotly_chart, style_fig(fig, height=360, title=f"향후 {CHART_DAYS}일 재고 흐름 (리드타임 {p_lt}일)"))

    with st.expander("📅 일자별 이벤트 로그 / 입고 스케줄"):
        if sim["events"]:
            wide(st.dataframe, pd.DataFrame(sim["events"], columns=["일자", "내용", "평택 잔여", "VF 잔여"]).round(0),
                 hide_index=True)
        else:
            st.write("기간 내 이벤트가 없습니다.")
        if it.inbound_detail:
            st.write("**공장 입고 스케줄** (발주일이 기준일 당일 또는 이전이면 발주완료로 반영)")
            wide(st.dataframe, pd.DataFrame([{
                "품목군": x["group"], "발주일": x["order"], "입고일": x["arrival"], "수량": int(x["qty"]),
                "날짜": "추정" if x["est"] else "확정",
                "상태": "✅ 발주완료 (반영)" if x["ordered"] else "🕓 미발주 계획 (미반영)"} for x in it.inbound_detail]),
                hide_index=True)


# ─────────────────────────────── 홈 ───────────────────────────────
def render_home(ctx):
    base, fcs, vat, inv = ctx["base"], ctx["fcs"], ctx["vat_incl"], ctx["inv_df"]
    cur = (base.year, base.month)
    k = 1.0 if vat else 1 / 1.1
    if ctx["stale_days"] >= 2:
        st.warning(f"엑셀 재고 기준일이 {base:%m/%d}로 오늘보다 {ctx['stale_days']}일 전입니다. "
                   "최신 재고·PO가 반영된 파일인지 확인하세요.")
    for msg in ctx["home_alerts"]:
        st.info(msg)

    r_act, r_conf, r_pred = (month_tot(fcs, cur, vat, kk) for kk in ("실적", "확정", "예측"))
    r_all = r_act + r_conf + r_pred
    r_dem = month_tot(fcs, cur, vat, demand=True)
    tgt = total_target(ctx["targets"], cur)
    r_ly = month_tot(fcs, (cur[0] - 1, cur[1]), vat, "실적")
    risk = sum(e.get("risk", 0) * e["price"] * k for f in fcs for e in f["events"] if e["kind"] == "확정")
    n_now = int((inv["_tier"] <= 1).sum())
    n_soon = int(inv["_tier"].isin([2, 3]).sum())

    c = st.columns(4)
    c[0].metric(f"💰 {cur[1]}월 예상 매출", won(r_all) if fcs else "-",
                delta=(f"목표 대비 {r_all / tgt:.0%}" if tgt else f"전년 동월 대비 {(r_all / r_ly - 1) * 100:+.0f}%" if r_ly else None),
                delta_color="off" if tgt else "normal")
    c[1].metric("✅ 확보 매출 (실적+확정)", won(r_act + r_conf) if fcs else "-",
                help=f"실적 {won(r_act)} · 확정 {won(r_conf)} · 예측 {won(r_pred)}")
    c[2].metric("🏭 지금 공장 발주", f"{n_now}개 SKU", delta=f"14일 내 {n_soon}개 더", delta_color="off")
    c[3].metric("⛔ 재고부족 손실 (이번 달)", won(r_dem - r_all) if fcs else "-",
                delta=(f"확정 PO 납품 위험 {won(risk)}" if risk else None), delta_color="inverse")
    if tgt:
        show_progress(r_act + r_conf, tgt, f"{cur[1]}월 확보")
        show_progress(r_all, tgt, f"{cur[1]}월 예상")

    st.write("")
    a, b, cc = st.columns(3)
    with a:
        with box():
            sec("🏭 공장 발주", "데드라인이 가까운 순")
            todo = inv[inv["_tier"].isin([0, 1, 2, 3])]
            if len(todo):
                wide(st.dataframe, todo[["상태", "품명", "발주 데드라인", "권장 발주량"]], hide_index=True,
                     column_config={"발주 데드라인": date_col("데드라인"), "권장 발주량": num_col("권장량")},
                     height=min(36 * len(todo) + 40, 320))
            else:
                st.success("14일 안에 발주할 품목이 없습니다.")
            st.caption("자세한 내용은 '재고 · 공장 발주' 탭")
    with b:
        with box():
            sec("🚚 입고 · 납품 점검", "신규 발주가 아니라 일정 조정이 필요한 것")
            rows = [{"구분": "입고 전 결품", "품명": r["품명"], "내용": r["상세"].replace("🟠 ", "")}
                    for _, r in inv[inv["_tier"] == 1.5].iterrows()]
            for f in fcs:
                es = [e for e in f["events"] if e["kind"] == "확정" and e.get("risk", 0) > 1e-9]
                if es:
                    amt = sum(e["risk"] * e["price"] * k for e in es)
                    rows.append({"구분": "확정 PO 재고 없음", "품명": f["name"],
                                 "내용": f"{es[0]['date']:%m/%d}~ {sum(e['risk'] for e in es):,.0f}개 모자람 ({won(amt)})"})
            if rows:
                wide(st.dataframe, pd.DataFrame(rows), hide_index=True, height=min(36 * len(rows) + 40, 320))
            else:
                st.success("점검할 항목이 없습니다.")
    with cc:
        with box():
            sec("🛒 쿠팡 PO 일정", "지연 PO + 앞으로 7일")
            rows = [{"일자": f["next"], "품명": f["name"], "구분": "⏰ 지연", "예상 금액": None} for f in fcs if f["overdue"]]
            for f in fcs:
                if f.get("manual"):
                    continue
                for e in f["events"]:
                    if base < e["date"] <= base + timedelta(days=7) and e["kind"] != "실적":
                        rows.append({"일자": e["date"], "품명": f["name"], "구분": e["kind"],
                                     "예상 금액": int(ev_rev(e, vat))})
            if rows:
                pdf = pd.DataFrame(rows)
                pdf = (pdf.groupby(["일자", "품명", "구분"], as_index=False, dropna=False)["예상 금액"].sum(min_count=1))
                pdf = pdf[(pdf["구분"] == "⏰ 지연") | (pdf["예상 금액"].fillna(0) > 0)]   # 재고가 없어 0원인 예측 PO는 숨김
                pdf = pdf.assign(_o=(pdf["구분"] != "⏰ 지연").astype(int)).sort_values(["_o", "일자"]).drop(columns="_o")
                wide(st.dataframe, pdf, hide_index=True, height=min(36 * len(pdf) + 40, 320),
                     column_config={"일자": date_col("일자"), "예상 금액": num_col("예상 금액")})
            else:
                st.write("예정된 PO가 없습니다.")

    if fcs:
        st.write("")
        l, r = st.columns([1, 1])
        with l:
            with box():
                df = group_rows(fcs, cur, ctx)
                fig = go.Figure()
                y = list(df["상품군"])[::-1]
                for kk in ("실적", "확정", "예측"):
                    fig.add_trace(go.Bar(y=y, x=list(df[kk])[::-1], orientation="h", name=kk,
                                         marker_color=C[{"실적": "act", "확정": "conf", "예측": "pred"}[kk]],
                                         hovertemplate="%{x:,.0f}원"))
                if df["목표"].notna().any():
                    fig.add_trace(go.Scatter(y=y, x=list(df["목표"])[::-1], mode="markers", name="목표",
                                             marker=dict(symbol="line-ns-open", size=22, color="#222", line=dict(width=3))))
                fig.update_layout(barmode="stack")
                wide(st.plotly_chart, style_fig(fig, height=max(240, 48 * len(df) + 110), title=f"상품군별 {cur[1]}월 매출"))
        with r:
            with box():
                cm = [add_months(cur[0], cur[1], i) for i in range(-5, FORECAST_MONTHS)]
                labels = [f"{a_ % 100:02d}.{b_:02d}" for a_, b_ in cm]
                fig = go.Figure()
                for kk in ("실적", "확정", "예측"):
                    fig.add_trace(go.Bar(x=labels, y=[month_tot(fcs, mm, vat, kk) for mm in cm], name=kk,
                                         marker_color=C[{"실적": "act", "확정": "conf", "예측": "pred"}[kk]],
                                         hovertemplate="%{y:,.0f}원"))
                fig.add_trace(go.Scatter(x=labels, y=[month_tot(fcs, (mm[0] - 1, mm[1]), vat, "실적") for mm in cm],
                                         name="전년 동월", mode="lines+markers", line=dict(color=C["ly"], dash="dot")))
                fig.update_layout(barmode="stack", hovermode="x unified")
                wide(st.plotly_chart, style_fig(fig, height=max(240, 48 * len(df) + 110), title="월별 매출 (최근 6개월 + 예측)"))


# ─────────────────────────────── 공장 발주 시뮬레이터 ───────────────────────────────
def scenario_metrics(fs, base, vat, month):
    k = 1.0 if vat else 1 / 1.1
    fut = [e for f in fs for e in f["events"] if e["date"] > base]
    return dict(
        rev=sum(ev_rev(e, vat) for e in fut),
        rev_m=sum(ev_rev(e, vat) for e in fut if mkey(e["date"]) == month),
        loss=sum((e["demand"] - e["deliv"]) * e["price"] * k for e in fut),
        risk=sum(e.get("risk", 0) * e["price"] * k for e in fut if e["kind"] == "확정"))


def render_simulator(ctx):
    base, items, inv = ctx["base"], ctx["items"], ctx["inv_df"]
    byname = {i.name: i for i in items}
    vat, cfg = ctx["vat_incl"], ctx["po_cfg"]
    cur = (base.year, base.month)
    sec("🏭 공장 발주 시뮬레이터",
        "발주할 품목·수량·입고일을 넣으면 결품일, 발주 상태, 회복되는 매출이 바로 계산됩니다. 여러 품목을 한꺼번에 넣을 수 있습니다.")

    need = inv[inv["_tier"].isin([0, 1, 2, 3])].head(5)
    seed = [{"품목": r["품명"], "발주일": base, "입고일": None,
             "수량": int(r["권장 발주량"] or byname[r["품명"]].moq)} for _, r in need.iterrows()]
    if not seed:
        it0 = items[0]
        seed = [{"품목": it0.name, "발주일": base, "입고일": None, "수량": int(it0.moq)}]
    edited = st.data_editor(
        pd.DataFrame(seed), num_rows="dynamic", hide_index=True, key=f"sim_{ctx['sig']}_{base}",
        column_config={
            "품목": st.column_config.SelectboxColumn("품목", options=sorted(byname), required=True, width="large"),
            "발주일": st.column_config.DateColumn("발주일", format="YYYY-MM-DD"),
            "입고일": st.column_config.DateColumn("입고일", format="YYYY-MM-DD", help="비우면 발주일 + 품목 리드타임"),
            "수량": st.column_config.NumberColumn("수량(개)", min_value=0, step=1, format="%d")})
    st.caption("기본값은 지금 발주가 필요한 품목과 권장 발주량입니다. 행을 추가·삭제하거나 숫자를 바꿔보세요.")

    orders, warns = {}, []
    for _, r in edited.iterrows():
        nm, q = r.get("품목"), to_num(r.get("수량"))
        if not isinstance(nm, str) or nm not in byname or not q or q <= 0:
            continue
        it = byname[nm]
        od = as_date(r.get("발주일")) or base
        ar = as_date(r.get("입고일")) or od + timedelta(days=it.lt)
        if ar < od + timedelta(days=it.lt):
            warns.append(f"{nm}: 입고일 {ar:%m/%d}이 리드타임({it.lt}일) 기준 가장 빠른 입고일({od + timedelta(days=it.lt):%m/%d})보다 빠릅니다.")
        if ar < base:
            warns.append(f"{nm}: 입고일이 기준일보다 앞이라 제외했습니다.")
            continue
        orders.setdefault(nm, []).append((ar, q, od))
    for w in warns:
        st.warning(w)
    if not orders:
        st.info("품목과 수량을 입력하면 결과가 나옵니다.")
        return

    res, detail = [], {}
    for nm, lst in orders.items():
        it = byname[nm]
        it2 = dataclasses.replace(it, inbound=sorted(it.inbound + [(a, q, False) for a, q, _ in lst]))
        sim0, ev0 = ctx["calc"][nm]
        sim1 = simulate(it2, base, it.adu, ctx["trig_mode"], ctx["trig_value"], ctx["po_mult"], ctx["transfer_days"])
        ev1 = evaluate(it2, sim1, base, it.adu, it.safety, it.lt, ctx["target_days"])
        f0 = [f for f in ctx["fcs"] if f["item"] is not None and f["item"].name == nm]
        f1 = [forecast_po(f["name"], ctx["po_by_name"][f["name"]], base, ctx["horizon_end"], it2, cfg) for f in f0]
        m0, m1 = scenario_metrics(f0, base, vat, cur), scenario_metrics(f1, base, vat, cur)
        so = lambda s: (base + timedelta(days=s["stockout"])) if s["stockout"] is not None else None
        res.append({
            "품목": nm, "추가 입고": ", ".join(f"{a:%m/%d} {q:,.0f}개" for a, q, _ in sorted(lst)),
            "상태 (전)": badge(ev0), "상태 (후)": badge(ev1),
            "평택 결품일 (전)": so(sim0), "평택 결품일 (후)": so(sim1),
            "다음 발주 데드라인 (후)": ev1["deadline"],
            "회복 매출": int(m1["rev"] - m0["rev"]), f"{cur[1]}월 매출 증가": int(m1["rev_m"] - m0["rev_m"]),
            "확정 PO 위험 해소": int(m0["risk"] - m1["risk"]), "남은 재고부족 손실": int(m1["loss"]),
            "PO 매칭": "O" if f0 else "X",
        })
        detail[nm] = (it, it2, sim0, sim1, f0, f1)
    rdf = pd.DataFrame(res)

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("📈 회복 매출 (예측 기간)", won(rdf["회복 매출"].sum()), help=f"~{ctx['horizon_end']:%Y-%m}")
    k2.metric(f"📅 {cur[1]}월 매출 증가", won(rdf[f"{cur[1]}월 매출 증가"].sum()))
    k3.metric("🚚 확정 PO 위험 해소", won(rdf["확정 PO 위험 해소"].sum()))
    k4.metric("⛔ 그래도 남는 손실", won(rdf["남은 재고부족 손실"].sum()))
    if not cfg["cap"]:
        st.info("사이드바에서 '재고 제약 반영'이 꺼져 있어 매출 변화가 0으로 나옵니다. 켜고 보세요.")
    ccfg = {c_: num_col(c_) for c_ in ("회복 매출", f"{cur[1]}월 매출 증가", "확정 PO 위험 해소", "남은 재고부족 손실")}
    ccfg.update({c_: date_col(c_) for c_ in ("평택 결품일 (전)", "평택 결품일 (후)", "다음 발주 데드라인 (후)")})
    wide(st.dataframe, rdf, hide_index=True, column_config=ccfg)
    st.caption("회복 매출 = 추가 입고 덕분에 납품 가능해지는 예측 PO 금액. PO 매칭 X = 쿠팡PO 탭과 연결이 안 돼 매출 변화를 계산할 수 없는 품목.")

    pick = st.selectbox("그래프로 볼 품목", list(detail), key="sim_pick")
    it, it2, sim0, sim1, f0, f1 = detail[pick]
    l, r = st.columns(2)
    n = 120
    days = [base + timedelta(days=i) for i in range(n)]
    with l:
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=days, y=sim0["main"][:n], name="발주 전", line=dict(color=C["before"], width=2, dash="dash")))
        fig.add_trace(go.Scatter(x=days, y=sim1["main"][:n], name="발주 후", line=dict(color=C["after"], width=3)))
        for a, q, _ in orders[pick]:
            if a < days[-1]:
                fig.add_vline(x=a.isoformat(), line_dash="dot", line_color=C["after"])
        if it.safety > 0:
            fig.add_hline(y=it.safety, line_dash="dashdot", line_color=C["safe"], annotation_text="안전재고")
        fig.update_layout(hovermode="x unified", yaxis_title="평택 재고 (개)")
        wide(st.plotly_chart, style_fig(fig, height=340, title="평택 재고 (발주 전 vs 후)"))
    with r:
        cm = [add_months(cur[0], cur[1], i) for i in range(FORECAST_MONTHS)]
        labels = [f"{a_ % 100:02d}.{b_:02d}" for a_, b_ in cm]
        fut = lambda fs, mm: sum(ev_rev(e, vat) for f in fs for e in f["events"] if e["date"] > base and mkey(e["date"]) == mm)
        fig = go.Figure()
        fig.add_trace(go.Bar(x=labels, y=[fut(f0, mm) for mm in cm], name="발주 전", marker_color=C["before"], hovertemplate="%{y:,.0f}원"))
        fig.add_trace(go.Bar(x=labels, y=[fut(f1, mm) for mm in cm], name="발주 후", marker_color=C["after"], hovertemplate="%{y:,.0f}원"))
        fig.update_layout(barmode="group", hovermode="x unified")
        wide(st.plotly_chart, style_fig(fig, height=340, title="기준일 이후 월별 PO 매출 (전 vs 후)"))


# ─────────────────────────────── 엑셀 리포트 ───────────────────────────────
def build_report(ctx) -> bytes:
    import io
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    base, fcs, vat = ctx["base"], ctx["fcs"], ctx["vat_incl"]
    k = 1.0 if vat else 1 / 1.1
    cur = (base.year, base.month)
    cm = [add_months(cur[0], cur[1], i) for i in range(-3, FORECAST_MONTHS)]
    mlab = lambda t: f"{t[0]}-{t[1]:02d}"
    inv = ctx["inv_df"].drop(columns=["_tier", "_dpol"])
    out = {}

    r_all = month_tot(fcs, cur, vat)
    tgt = total_target(ctx["targets"], cur)
    out["요약"] = pd.DataFrame([
        ("기준일", f"{base:%Y-%m-%d}"), ("작성일", f"{today_kst():%Y-%m-%d}"), ("금액 기준", "VAT 포함" if vat else "VAT 제외"),
        (f"{cur[1]}월 실적", int(month_tot(fcs, cur, vat, "실적"))), (f"{cur[1]}월 확정", int(month_tot(fcs, cur, vat, "확정"))),
        (f"{cur[1]}월 예측", int(month_tot(fcs, cur, vat, "예측"))), (f"{cur[1]}월 예상 합계", int(r_all)),
        (f"{cur[1]}월 목표", int(tgt) if tgt else "-"), ("목표 대비", f"{r_all / tgt:.0%}" if tgt else "-"),
        ("지금 공장 발주 SKU", int((ctx["inv_df"]["_tier"] <= 1).sum())),
        ("14일 내 발주 SKU", int(ctx["inv_df"]["_tier"].isin([2, 3]).sum())),
    ], columns=["항목", "값"])
    out["공장발주_필요"] = inv[ctx["inv_df"]["_tier"] <= 3]
    out["재고_전체"] = inv
    if fcs:
        rows = []
        for f in fcs:
            row = {"상품군": f["group"], "품명": f["name"], "차기 PO": f["next"]}
            for mm in cm:
                row[mlab(mm)] = int(month_tot([f], mm, vat))
            rows.append(row)
        out["월별매출_SKU"] = pd.DataFrame(rows).sort_values(["상품군", mlab(cur)], ascending=[True, False])
        g = []
        for grp in sorted({f["group"] for f in fcs}):
            fs = [f for f in fcs if f["group"] == grp]
            row = {"상품군": grp}
            for mm in cm:
                row[mlab(mm)] = int(month_tot(fs, mm, vat))
                t = ctx["targets"].get(mm, {}).get(grp)
                if t:
                    row[mlab(mm) + " 목표"] = int(t)
            g.append(row)
        out["상품군별"] = pd.DataFrame(g)
        loss = []
        for f in fcs:
            fut = [e for e in f["events"] if e["date"] > base]
            ls = sum((e["demand"] - e["deliv"]) * e["price"] * k for e in fut)
            rk = sum(e.get("risk", 0) * e["price"] * k for e in fut)
            if ls > 0 or rk > 0:
                loss.append({"상품군": f["group"], "품명": f["name"], "재고부족 손실": int(ls), "확정 PO 납품 위험": int(rk),
                             "첫 미납일": min((e["date"] for e in fut if e["demand"] - e["deliv"] > 1e-9 or e.get("risk", 0) > 1e-9), default=None)})
        out["재고부족_손실"] = pd.DataFrame(loss).sort_values("재고부족 손실", ascending=False) if loss else pd.DataFrame(columns=["품명"])
        ev_rows = [{"일자": e["date"], "구분": e["kind"], "상품군": f["group"], "품명": f["name"], "PO 수량": int(e["demand"]),
                    "납품(예상)": int(e["deliv"]), "단가": int(e["price"]), "매출": int(ev_rev(e, vat)),
                    "재고 모자람": int(e.get("risk", 0)) or None}
                   for f in fcs for e in f["events"] if e["date"] > base]
        out["PO일정"] = pd.DataFrame(ev_rows).sort_values(["일자", "품명"]) if ev_rows else pd.DataFrame(columns=["일자"])

    buf = io.BytesIO()
    head_fill, head_font = PatternFill("solid", fgColor="1F4E79"), Font(bold=True, color="FFFFFF")
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for name, d in out.items():
            d.to_excel(xw, sheet_name=name, index=False)
            ws = xw.sheets[name]
            ws.freeze_panes = "A2"
            for cell in ws[1]:
                cell.fill, cell.font = head_fill, head_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            for j, col in enumerate(d.columns, start=1):
                vals = [str(col)] + [str(v) for v in d[col].head(200).tolist()]
                ws.column_dimensions[get_column_letter(j)].width = min(max(len(v) for v in vals) * 1.3 + 2, 60)
                is_num = pd.api.types.is_numeric_dtype(d[col])
                is_dt = d[col].map(lambda v: isinstance(v, date)).any() if len(d) else False
                for cell in ws.iter_rows(min_row=2, min_col=j, max_col=j):
                    c0 = cell[0]
                    if is_num:
                        c0.number_format = "#,##0"
                    elif is_dt:
                        c0.number_format = "yyyy-mm-dd"
    return buf.getvalue()


def manual_template() -> bytes:
    import io
    df = pd.DataFrame([{"날짜": today_kst(), "상품군": EXTRA_TARGET_GROUPS[0], "매출금액": 1000000, "메모": "예시 행 (지우고 쓰세요)"}])
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="기타매출", index=False)
        ws = xw.sheets["기타매출"]
        for col, w in zip("ABCD", (14, 18, 14, 24)):
            ws.column_dimensions[col].width = w
        ws["A2"].number_format = "yyyy-mm-dd"
        ws["C2"].number_format = "#,##0"
    return buf.getvalue()


def target_template(groups, months) -> bytes:
    import io
    df = pd.DataFrame({"상품군": ["전체"] + list(groups), **{f"{y}-{m:02d}": [None] * (len(groups) + 1) for y, m in months}})
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="목표", index=False)
        ws = xw.sheets["목표"]
        ws.column_dimensions["A"].width = 18
        ws.cell(row=len(groups) + 4, column=1, value="단위: 만원. '전체'를 비우면 상품군 합계를 전체 목표로 씁니다.")
    return buf.getvalue()


# ─────────────────────────────── 데이터 점검 ───────────────────────────────
def build_notes(ctx):
    base, items, sheets, wb = ctx["base"], ctx["items"], ctx["sheets"], ctx["wb"]
    po_recs, po_names, mapping = ctx["po_recs"], ctx["po_names"], ctx["mapping"]
    notes = []
    if ctx["overrides"]:
        notes.append(f"입고 일정 {len(ctx['overrides'])}개 컬럼을 화면에서 직접 보정했습니다 (보정한 날짜는 확정으로 계산).")
    not_yet = [(it, x) for it in items for x in it.inbound_detail if not x["ordered"]]
    if not_yet:
        cols_ = {(x["group"], x["order"], x["arrival"]) for _, x in not_yet}
        notes.append(f"발주일이 기준일({base:%m/%d})보다 뒤인 입고 계획 {len(cols_)}개 컬럼(품목 {len({it.name for it, _ in not_yet})}개)은 "
                     "아직 발주 전으로 보고 재고 계산에서 제외했습니다.")
    if any(x["order"] is None and x["arrival"] >= base for it in items for x in it.inbound_detail):
        notes.append("발주일이 적혀 있지 않은 입고 컬럼은 발주 완료로 간주했습니다.")
    est_cnt = sum(1 for it in items for _, _, e in it.inbound if e)
    if est_cnt:
        notes.append(f"입고일이 '11월초/12월 중순' 같은 표현인 {est_cnt}건은 순의 끝날(초=10일, 중순=20일, 말=말일)로 보수적으로 추정했습니다.")
    for sh in sheets:
        if sh["unparsed"]:
            notes.append(f"[{sh['name']}] 날짜를 해석하지 못한 입고 컬럼: {', '.join(sh['unparsed'])}")
    if ctx["excluded"]:
        notes.append(f"최신 시트({ctx['latest']})에 없어 제외한 품목 {len(ctx['excluded'])}개: {', '.join(ctx['excluded'])}")
    planned = [i for i in items if i.plan]
    if planned:
        notes.append(f"기준일 이후 날짜에 이미 입력된 출고 수량({len(planned)}개 품목, ~{max(i.plan_until for i in planned):%m/%d})은 "
                     "쿠팡 PO 예정으로 보고 그대로 사용했습니다.")
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
    if wb.get("manual"):
        m = wb["manual"]
        notes.append(f"엑셀 '기타매출' 시트에서 {len(m)}행, 합계 {won(sum(x['amount'] for x in m))}을 읽었습니다 "
                     f"(상품군: {', '.join(sorted({x['group'] for x in m}))}).")
    if wb.get("targets"):
        notes.append(f"엑셀 '목표' 시트에서 {len(wb['targets'])}개월 목표를 읽었습니다.")
    return notes


# ─────────────────────────────── 메인 ───────────────────────────────
def main():
    st.set_page_config(layout="wide", page_title="쿠팡 VF & 평택 발주·매출 대시보드", page_icon="📦")
    st.markdown(CSS, unsafe_allow_html=True)

    # ── 사이드바: 기본
    st.sidebar.markdown("### 📁 데이터")
    up = st.sidebar.file_uploader("통합 마스터 엑셀 (.xlsx)", type=["xlsx"], label_visibility="collapsed")
    st.title("📦 쿠팡 VF & 평택 발주 · 매출 대시보드")
    if not up:
        st.info("👈 왼쪽 사이드바에서 통합 마스터 엑셀을 업로드해 주세요. "
                "('쿠팡PO' 이력 탭 + 월별 재고 시트, 선택적으로 '목표' 시트)")
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

    base = st.sidebar.date_input("재고 기준일", snapshot or today_kst(), key=f"base_{sig}",
                                 help="재고 수량이 입력된 날짜. 기본값은 엑셀 헤더의 기준일(=TODAY() 저장값).")
    st.sidebar.markdown("### 💰 매출 보기")
    with st.sidebar:
        vat_incl = toggle("VAT 포함", True, key="vat", help="PO 탭 단가가 VAT 포함 기준입니다. 끄면 ÷1.1 공급가.")
        po_cap = toggle("재고 제약 반영", True, key="cap",
                        help="예측 PO를 평택 재고 + 발주완료 입고분 안에서만 납품하는 것으로 계산합니다.")
    po_mult_pct = st.sidebar.slider("수요 보정 (%)", 50, 200, 100, step=5, key="mult",
                                    help="시즌·행사로 출고가 늘거나 줄 것으로 보면 조정. 예측 PO 간격이 그만큼 짧아/길어집니다.")
    report_slot = st.sidebar.container()

    with st.sidebar.expander("⚙️ 고급 설정 · 재고 시뮬레이션"):
        win_std = st.number_input("출고 평균 기간 – 일반(일)", 1, 60, 7)
        win_mr = st.number_input("출고 평균 기간 – 밀크런(일)", 1, 60, 28,
                                 help="밀크런은 PO 단위로 몰아서 나가서 7일 평균은 0이 되기 쉽습니다.")
        mr_kw = st.text_input("밀크런 구분 키워드", "밀크런")
        default_lt = st.number_input("기본 리드타임 (일)", 1, 120, 35)
        use_hist_lt = st.checkbox("발주→입고 이력으로 리드타임 자동 추정", True)
        trig_mode_label = st.radio("쿠팡 PO 트리거 방식", ["VF 잔여일수 기준", "고정 수량 기준"])
        trig_mode = "days" if trig_mode_label.startswith("VF") else "units"
        if trig_mode == "days":
            trig_value = st.number_input("VF 잔여 며칠분 밑이면 PO?", 1, 60, 10)
        else:
            trig_value = st.number_input("VF 잔여 수량 트리거 (개)", 0, 100000, 200, step=10)
        po_mult = st.selectbox("쿠팡 PO 배수", [1, 2, 3], format_func=lambda x: f"MOQ × {x}")
        transfer_days = st.number_input("평택→VF 이동 소요일", 0, 14, 0)
        target_days = st.number_input("권장 발주량 산정 기간 (일)", 7, 120, 45)
        use_plan = st.checkbox("시트에 미리 입력된 미래 출고 반영", True)
        include_stale = st.checkbox("최신 월 시트에 없는 품목도 포함", False)
    with st.sidebar.expander("⚙️ 고급 설정 · PO 예측"):
        po_mode_label = st.radio("출고율 기준", ["VF형=시트 최근 출고 / 밀크런=PO 실적", "모두 PO 실적 기준"])
        po_win_sheet = st.number_input("시트 출고 평균 기간 (일)", 1, 60, 7)
        po_win_po = st.number_input("PO 실적 집계 기간 (일)", 7, 120, 28)
        po_q_n = st.number_input("PO 1회 수량 = 최근 N회 중앙값, N", 1, 30, 8)
        po_weekend = st.checkbox("토·일 예측 PO는 월요일로", True)

    stale = (today_kst() - base).days
    st.markdown(f"<div class='kicker'>재고 기준일 {base:%Y-%m-%d} · 파일 {up.name} · PO {len(po_recs):,}행"
                + (f" · 마지막 PO {max(r['date'] for r in po_recs):%m/%d}" if po_recs else "") + "</div>",
                unsafe_allow_html=True)

    t_home, t_rev, t_inv, t_sim, t_data = st.tabs(
        ["🏠 홈", "💰 매출 · 쿠팡 PO", "📦 재고 · 공장 발주", "🏭 발주 시뮬레이터", "🛠️ 데이터 설정"])

    # ── 데이터 설정 ①: 입고 일정 보정 (품목 계산 전에 반영)
    latest_sheet = sorted(sheets, key=lambda s_: (s_["year"], s_["month"]))[-1]
    overrides = {}
    with t_data:
        sec("🗓️ 입고 일정 보정", "엑셀에 '11월초'처럼 대략 적힌 날짜를 실제 날짜로 바꾸면 확정 날짜로 계산합니다. "
            "새 파일을 올리면 초기화되니, 계속 쓰려면 엑셀 헤더에도 날짜를 적어 두세요.")
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
        else:
            st.caption("최신 월 시트에 입고 스케줄 컬럼이 없습니다.")
    sheets = apply_inbound_overrides(sheets, latest_sheet["name"], overrides)

    items, latest, entered, lt_est = build_items(
        sheets, base, win_std, win_mr, default_lt, use_hist_lt, mr_kw, include_stale, use_plan)
    if not items:
        st.error("표시할 품목이 없습니다.")
        return
    excluded = sorted({r["name"] for s in sheets for r in s["rows"]} - {i.name for i in items})

    # ── 데이터 설정 ②: PO 매칭 · 상품군 · 목표
    po_names = sorted({r["name"] for r in po_recs})
    auto = auto_match(po_names, [i.name for i in items])
    mapping = {p: v[0] for p, v in auto.items() if v[0]}
    with t_data:
        st.divider()
        if po_names:
            sec("🔗 PO 품명 ↔ 재고시트 품명 매칭",
                "매칭돼야 시트 출고율·평택 재고가 PO 예측에 쓰입니다. '유사'로 잡힌 건 맞는지 확인하세요.")
            mtbl = pd.DataFrame([{"PO 품명": p, "자동 매칭": auto[p][1], "재고시트 품명": auto[p][0] or NO_MATCH}
                                 for p in po_names])
            medit = st.data_editor(
                mtbl, hide_index=True, key=f"po_map_{sig}", disabled=["PO 품명", "자동 매칭"],
                column_config={"재고시트 품명": st.column_config.SelectboxColumn(
                    "재고시트 품명", options=[NO_MATCH] + sorted(i.name for i in items), required=True)})
            mapping = {r_["PO 품명"]: r_["재고시트 품명"] for _, r_ in medit.iterrows() if r_["재고시트 품명"] != NO_MATCH}
        st.divider()
        sec("🏷️ 상품군 기준", "품명에 키워드가 들어 있으면 그 상품군으로 묶습니다. 위에서부터 먼저 맞는 규칙을 씁니다.")
        gedit = st.data_editor(pd.DataFrame(DEFAULT_GROUP_RULES, columns=["키워드", "상품군"]),
                               num_rows="dynamic", hide_index=True, key="group_rules")
        rules = [(str(r_["키워드"]), str(r_["상품군"])) for _, r_ in gedit.iterrows()
                 if clean_txt(r_["키워드"]) and clean_txt(r_["상품군"])]

    hy, hm = add_months(base.year, base.month, FORECAST_MONTHS - 1)
    horizon_end = month_end(hy, hm)
    po_cfg = dict(mode="sheet" if po_mode_label.startswith("VF") else "po", win_sheet=int(po_win_sheet),
                  win_po=int(po_win_po), q_n=int(po_q_n), mult=po_mult_pct / 100, cap=po_cap,
                  weekend=po_weekend, entered=entered)
    fcs = forecast_all(po_recs, items, mapping, base, horizon_end, po_cfg) if po_recs else []
    for f in fcs:
        f["group"] = group_of(f["name"], rules)
    fcs += manual_entries(wb.get("manual", []), base)
    item_group = {i.name: group_of(i.name, rules) for i in items}
    groups = sorted(set(item_group.values()) | {f["group"] for f in fcs})

    tmonths = [add_months(base.year, base.month, i) for i in range(FORECAST_MONTHS)]
    with t_data:
        st.divider()
        sec("🎯 월 매출 목표 (만원)", "여기서 입력한 값은 이 화면에서만 유지됩니다. 팀이 같이 쓰려면 템플릿을 받아 "
            "마스터 엑셀에 '목표' 시트로 넣어 두세요 (다음 업로드부터 자동으로 채워짐).")
        xt = wb.get("targets", {})
        tdf = pd.DataFrame([{"상품군": g, **{f"{y}-{m:02d}": (xt.get((y, m), {}).get(g)) for y, m in tmonths}}
                            for g in ["전체"] + list(dict.fromkeys(EXTRA_TARGET_GROUPS + groups
                                                                     + sorted({g_ for d in xt.values() for g_ in d} - {"전체"})))])
        tedit = st.data_editor(tdf, hide_index=True, key=f"targets_{sig}", disabled=["상품군"],
                               column_config={f"{y}-{m:02d}": st.column_config.NumberColumn(f"{y}-{m:02d}", min_value=0, format="%d")
                                              for y, m in tmonths})
        st.download_button("⬇️ 목표 시트 템플릿 (.xlsx)", target_template(list(dict.fromkeys(EXTRA_TARGET_GROUPS + groups)), tmonths), file_name="목표_템플릿.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.divider()
        sec("🧾 쿠팡 외 매출 직접 입력 ('기타매출' 시트)",
            "네이버 두부모래처럼 쿠팡 PO가 아닌 매출은 마스터 엑셀에 '기타매출' 시트를 만들어 [날짜 | 상품군 | 매출금액]만 적으면 "
            "실적으로 합산됩니다. 금액은 VAT 포함, 날짜 대신 '2026-10'처럼 월만 적어도 됩니다. 상품군을 비우면 '네이버 두부모래'로 봅니다.")
        st.download_button("⬇️ 기타매출 시트 템플릿 (.xlsx)", manual_template(), file_name="기타매출_템플릿.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    targets = {}
    for (y, m) in tmonths:
        for _, r_ in tedit.iterrows():
            v = to_num(r_[f"{y}-{m:02d}"])
            if v:
                targets.setdefault((y, m), {})[r_["상품군"]] = v * 1e4
    for ym, d in xt.items():                     # 화면 범위 밖 월은 엑셀 값 그대로
        if ym not in tmonths:
            targets[ym] = {g: v * 1e4 for g, v in d.items()}

    next_po = {}
    for f in fcs:
        if f["item"] is not None and f["next"]:
            n0 = next_po.get(f["item"].name)
            next_po[f["item"].name] = min(n0, f["next"]) if n0 else f["next"]

    calc = {}
    for it in items:
        sim = simulate(it, base, it.adu, trig_mode, trig_value, po_mult, transfer_days)
        calc[it.name] = (sim, evaluate(it, sim, base, it.adu, it.safety, it.lt, target_days))

    po_by_name = {}
    for r in po_recs:
        po_by_name.setdefault(r["name"], []).append(r)
    ctx = dict(base=base, items=items, calc=calc, fcs=fcs, vat_incl=vat_incl, po_cfg=po_cfg, horizon_end=horizon_end,
               targets=targets, item_group=item_group, next_po=next_po, sig=sig, trig_mode=trig_mode,
               trig_value=trig_value, po_mult=po_mult, transfer_days=transfer_days, target_days=target_days,
               po_by_name=po_by_name, deadlines={n: c_[1] for n, c_ in calc.items()},
               months=months_span(fcs, base, horizon_end) if fcs else [], stale_days=stale,
               sheets=sheets, wb=wb, po_recs=po_recs, po_names=po_names, mapping=mapping, overrides=overrides,
               excluded=excluded, latest=latest)
    ctx["inv_df"] = inventory_table(ctx)

    unmatched = [p for p in po_names if p not in mapping]
    fuzzy = [p for p, v in auto.items() if v[1].startswith("유사") and mapping.get(p) == v[0]]
    alerts = []
    if unmatched or fuzzy:
        alerts.append(f"PO 품명 매칭 확인 필요: 미매칭 {len(unmatched)}개, 유사 매칭 {len(fuzzy)}개 → '🛠️ 데이터 설정' 탭")
    ctx["home_alerts"] = alerts

    with report_slot:
        try:
            st.download_button("⬇️ 엑셀 리포트 받기", build_report(ctx), file_name=f"발주매출_리포트_{base:%m%d}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               help="공장발주 필요 · 재고 전체 · 월별 매출(SKU/상품군) · 재고부족 손실 · PO 일정", **({"type": "primary"}))
        except Exception as e:
            st.caption(f"리포트 생성 실패: {e}")

    with t_home:
        render_home(ctx)
    with t_rev:
        render_po_tab(ctx)
    with t_inv:
        render_inventory(ctx)
    with t_sim:
        render_simulator(ctx)
    with t_data:
        st.divider()
        notes = build_notes(ctx)
        sec(f"🔎 데이터 점검 사항 ({len(notes)})")
        for n in notes:
            st.write("• " + n)
        if lt_est:
            st.write("**발주→입고 실적 리드타임 (품목군별 중앙값)**")
            wide(st.dataframe, pd.DataFrame([{"품목군": g, "중앙값(일)": v["median"], "표본수": v["n"], "정확한 날짜 쌍": v["n_exact"]}
                                             for g, v in lt_est.items()]), hide_index=True)


if __name__ == "__main__":
    main()
