import re
from pathlib import Path
from functools import lru_cache
import unicodedata

import openpyxl
import pandas as pd

from .config.exel_column_map import INVENTORY_COLUMN_MAP


# ======================
# 공통 유틸
# ======================

def _canon(s: object) -> str:
    """공백/괄호/구분기호 제거 + 소문자 정규화"""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = unicodedata.normalize("NFKC", str(s))
    return re.sub(r"\s+|\(|\)|\[|\]|{|}|/|\\|-|_", "", s).lower()


def _parse_korean_md(s: str, year: int):
    """'MM월 DD일' → Timestamp 변환"""
    if pd.isna(s):
        return pd.NaT
    if isinstance(s, pd.Timestamp):
        return s.normalize()
    txt = str(s).strip()
    m = re.match(r"(\d{1,2})\s*월\s*(\d{1,2})\s*일", txt)
    if m:
        try:
            return pd.Timestamp(year=year, month=int(m[1]), day=int(m[2]))
        except Exception:
            return pd.NaT
    return pd.to_datetime(txt, errors="coerce")


def _norm(s):
    return "" if pd.isna(s) else str(s).strip().lower()


def _is_return_text(s: object) -> bool:
    txt = _canon(s)
    return ("반품" in txt) or ("return" in txt)


# ======================
# 컬럼 선택기
# ======================

def _pick_date_col(df, col_date_lv0):
    for c in df.columns:
        if isinstance(c, tuple) and _norm(c[0]) == _norm(col_date_lv0):
            return c
        if not isinstance(c, tuple) and _norm(c) == _norm(col_date_lv0):
            return c
    raise KeyError(f"날짜 컬럼(level=0='{col_date_lv0}') 없음")


def _pick_arrival_col(df, col_arrival_lv0: str):
    for c in df.columns:
        if isinstance(c, tuple) and _norm(c[0]) == _norm(col_arrival_lv0):
            return c
        if not isinstance(c, tuple) and _norm(c) == _norm(col_arrival_lv0):
            return c
    return None


def _pick_product_col(df, col_prod_lv0):
    for c in df.columns:
        if isinstance(c, tuple) and _norm(c[0]) == _norm(col_prod_lv0):
            return c
        if not isinstance(c, tuple) and _norm(c) == _norm(col_prod_lv0):
            return c
    raise KeyError(f"품명 컬럼(level=0='{col_prod_lv0}') 없음")


def _pick_current_stock_col(df, col_stock_lv0):
    for c in df.columns:
        if isinstance(c, tuple) and _norm(c[0]) == _norm(col_stock_lv0):
            return c
        if not isinstance(c, tuple) and _norm(c) == _norm(col_stock_lv0):
            return c
    raise KeyError(f"현재고 컬럼(level=0='{col_stock_lv0}') 없음")


def _pick_optional_col(df, col_lv0: str | None):
    if not col_lv0:
        return None
    target = _norm(col_lv0)
    for c in df.columns:
        if isinstance(c, tuple):
            parts = [_norm(x) for x in c]
            if target in parts:
                return c
        elif _norm(c) == target:
            return c
    return None


def _pick_sales_qty_cols(df: pd.DataFrame):
    """납품/매출 블록의 수량(Kg) 컬럼"""
    chosen = []
    if isinstance(df.columns, pd.MultiIndex):
        for col in df.columns:
            parts = [_canon(x) for x in col]
            if not parts:
                continue
            top = parts[0]
            if (("납품" not in top) and ("매출" not in top)) or ("입고" in top) or ("수입" in top):
                continue
            rest = parts[1:]
            if any(("수량" in p) or ("kg" in p) for p in rest):
                chosen.append(col)
    else:
        for col in df.columns:
            p = _canon(col)
            if (("납품" in p) or ("매출" in p)) and not (("입고" in p) or ("수입" in p)):
                if ("수량" in p) or ("kg" in p):
                    chosen.append(col)
    if not chosen:
        raise ValueError("납품/매출 수량 열 없음")
    return chosen


def _pick_inbound_qty_cols(df: pd.DataFrame):
    chosen = []
    if isinstance(df.columns, pd.MultiIndex):
        for c in df.columns:
            parts = [_canon(x) for x in c]
            if not parts:
                continue
            if ("입고" in parts[0] or "수입" in parts[0]) and any(("수량" in p) or ("kg" in p) for p in parts[1:]):
                chosen.append(c)
    else:
        for c in df.columns:
            p = _canon(c)
            if (("입고" in p) or ("수입" in p)) and (("수량" in p) or ("kg" in p)):
                chosen.append(c)
    return chosen


def _pick_inbound_unit_price_cols(df: pd.DataFrame):
    chosen = []
    if isinstance(df.columns, pd.MultiIndex):
        for col in df.columns:
            parts = [_canon(x) for x in col]
            if ("입고" in parts[0] or "수입" in parts[0]) and any("단가" in p for p in parts[1:]):
                chosen.append(col)
    else:
        for col in df.columns:
            p = _canon(col)
            if ("입고" in p or "수입" in p) and "단가" in p:
                chosen.append(col)
    return chosen


def _pick_inbound_exchange_rate_cols(df: pd.DataFrame):
    chosen = []
    if isinstance(df.columns, pd.MultiIndex):
        for col in df.columns:
            parts = [_canon(x) for x in col]
            if ("입고" in parts[0] or "수입" in parts[0]) and any("환율" in p for p in parts[1:]):
                chosen.append(col)
    else:
        for col in df.columns:
            p = _canon(col)
            if ("입고" in p or "수입" in p) and "환율" in p:
                chosen.append(col)
    return chosen


# ======================
# 날짜/숫자 표준화
# ======================

def _make_event_date_series(df: pd.DataFrame, primary_col, arrival_col, year: int) -> pd.Series:
    """판매/출고 기준일: '날 짜' 우선, 없으면 도착일 폴백"""
    s1 = df[primary_col].apply(lambda x: _parse_korean_md(x, year))
    if arrival_col is None:
        return s1
    s2 = df[arrival_col].apply(lambda x: _parse_korean_md(x, year))
    return s1.where(~s1.isna(), s2)


def _make_inbound_date_series(df: pd.DataFrame, date_col, arrival_col, year: int) -> pd.Series:
    """입고 기준일: 도착일 우선, 없으면 '날 짜' 사용"""
    s1 = df[date_col].apply(lambda x: _parse_korean_md(x, year))
    if arrival_col is None:
        return s1
    s2 = df[arrival_col].apply(lambda x: _parse_korean_md(x, year))
    return s2.where(~s2.isna(), s1)


def _to_num_frame(frame: pd.DataFrame) -> pd.DataFrame:
    return frame.replace(r"[^\d\.\-]", "", regex=True).apply(pd.to_numeric, errors="coerce")


def _first_valid_by_row(frame: pd.DataFrame) -> pd.Series:
    """행 단위에서 첫 번째 유효 숫자 선택"""
    if frame.empty:
        return pd.Series(pd.NA, index=frame.index)
    num = _to_num_frame(frame)
    if num.shape[1] == 1:
        return num.iloc[:, 0]
    return num.bfill(axis=1).iloc[:, 0]


def _clean_text_series(s: pd.Series) -> pd.Series:
    out = s.where(~s.isna(), "").astype(str).str.strip()
    return out.replace({"nan": "", "None": "", "<NA>": ""})


def _extract_year_from_sheet_name(sheet_name: str) -> int | None:
    m = re.search(r"(19|20)\d{2}", str(sheet_name))
    if not m:
        return None
    try:
        return int(m.group(0))
    except Exception:
        return None


@lru_cache(maxsize=8)
def _available_inventory_years(file_path_str: str) -> tuple[int, ...]:
    file_path = Path(file_path_str)
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    years: set[int] = set()
    for s in wb.sheetnames:
        y = _extract_year_from_sheet_name(s)
        if y is not None:
            years.add(y)
    wb.close()
    return tuple(sorted(years))


def _years_to_scan_desc(file_path_str: str, target_year: int) -> tuple[int, ...]:
    years = [y for y in _available_inventory_years(file_path_str) if y <= int(target_year)]
    years.sort(reverse=True)
    return tuple(years)


# ======================
# 준비/캐시: 엑셀 1회 로딩 + 파생 컬럼 사전계산
# ======================

@lru_cache(maxsize=4)
def load_inventory_df(file_path_str: str, year: int) -> pd.DataFrame:
    """
    파일/연도 단위로 엑셀을 1회만 읽고, 필요한 파생 컬럼을 생성해 반환.
    """
    file_path = Path(file_path_str)
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    sheets = [s for s in wb.sheetnames if str(year) in s]
    wb.close()
    if not sheets:
        return pd.DataFrame()

    parts = []
    for sn in sheets:
        raw_part = pd.read_excel(
            file_path,
            sheet_name=sn,
            engine="openpyxl",
            header=[2, 3],
            dtype=str,
        )
        parts.append(raw_part)
    raw = pd.concat(parts, ignore_index=True) if len(parts) > 1 else parts[0]

    date_col = _pick_date_col(raw, INVENTORY_COLUMN_MAP["date"])
    prod_col = _pick_product_col(raw, INVENTORY_COLUMN_MAP["product_name"])
    arrival_col = _pick_arrival_col(raw, INVENTORY_COLUMN_MAP.get("arrival_date", "도착일"))
    stock_col = _pick_current_stock_col(raw, INVENTORY_COLUMN_MAP["current_stock"])
    supplier_col = _pick_optional_col(raw, INVENTORY_COLUMN_MAP.get("supplier"))
    customer_col = _pick_optional_col(raw, INVENTORY_COLUMN_MAP.get("customer"))
    note_col = _pick_optional_col(raw, INVENTORY_COLUMN_MAP.get("note"))

    out = pd.DataFrame(index=raw.index)
    out["__product__"] = _clean_text_series(raw[prod_col])
    out["__event_date__"] = _make_event_date_series(raw, date_col, arrival_col, year)
    out["__in_date__"] = _make_inbound_date_series(raw, date_col, arrival_col, year)
    out["__row_date__"] = raw[date_col].apply(lambda x: _parse_korean_md(x, year))
    out["__stock__"] = pd.to_numeric(raw[stock_col], errors="coerce")
    out["__supplier__"] = _clean_text_series(raw[supplier_col]) if supplier_col is not None else ""
    out["__customer__"] = _clean_text_series(raw[customer_col]) if customer_col is not None else ""
    out["__note__"] = _clean_text_series(raw[note_col]) if note_col is not None else ""

    # 판매 수량
    try:
        sales_qty_cols = _pick_sales_qty_cols(raw)
        out["__sales_qty__"] = _to_num_frame(raw[sales_qty_cols]).fillna(0.0).sum(axis=1)
    except Exception:
        out["__sales_qty__"] = 0.0

    # 입고 수량
    inbound_qty_cols = _pick_inbound_qty_cols(raw)
    if inbound_qty_cols:
        out["__in_qty__"] = _to_num_frame(raw[inbound_qty_cols]).fillna(0.0).sum(axis=1)
    else:
        out["__in_qty__"] = 0.0

    # 단가 / 환율은 행 기준 첫 유효값 채택
    price_cols = _pick_inbound_unit_price_cols(raw)
    rate_cols = _pick_inbound_exchange_rate_cols(raw)
    out["__unit_price__"] = _first_valid_by_row(raw[price_cols]) if price_cols else pd.NA
    out["__exchange_rate__"] = _first_valid_by_row(raw[rate_cols]) if rate_cols else pd.NA

    out["__event_date__"] = pd.to_datetime(out["__event_date__"], errors="coerce")
    out["__in_date__"] = pd.to_datetime(out["__in_date__"], errors="coerce")
    out["__row_date__"] = pd.to_datetime(out["__row_date__"], errors="coerce")
    out["__sheet_year__"] = int(year)

    return out


@lru_cache(maxsize=4)
def load_inventory_history_df(file_path_str: str, year: int) -> pd.DataFrame:
    """
    대상 연도 이하(예: 2025 -> 2025, 2024, 2023...) 시트를 모두 로딩해 결합.
    입고 역추적 시 과거 연도까지 확장하기 위한 데이터셋.
    """
    years_desc = _years_to_scan_desc(file_path_str, int(year))
    if not years_desc:
        return load_inventory_df(file_path_str, int(year))

    frames: list[pd.DataFrame] = []
    for yy in sorted(years_desc):
        d = load_inventory_df(file_path_str, int(yy))
        if not d.empty:
            frames.append(d)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def _load_inbound_events_progressive(
    file_path: Path,
    year: int,
    product: str,
    cutoff: pd.Timestamp,
    base_df: pd.DataFrame | None = None,
    min_required_qty: float | None = None,
) -> pd.DataFrame:
    """
    1) 요청 연도(예: 2025) 먼저 조회
    2) 입고가 없거나 수량이 부족할 때만 2024, 2023... 순으로 역추적
    """
    if base_df is None:
        base_df = load_inventory_df(str(file_path), year)

    frames: list[pd.DataFrame] = []
    base_ev = extract_inbound_events(file_path, year, product, cutoff=cutoff, df=base_df)
    if not base_ev.empty:
        frames.append(base_ev)

    qty_sum = float(pd.to_numeric(base_ev["quantity"], errors="coerce").fillna(0.0).sum()) if not base_ev.empty else 0.0
    needs_more = (qty_sum <= 1e-9) if min_required_qty is None else (qty_sum + 1e-9 < float(min_required_qty))

    if needs_more:
        for yy in _years_to_scan_desc(str(file_path), int(year)):
            if yy >= int(year):
                continue
            d_prev = load_inventory_df(str(file_path), int(yy))
            if d_prev.empty:
                continue

            prev_ev = extract_inbound_events(file_path, int(yy), product, cutoff=cutoff, df=d_prev)
            if prev_ev.empty:
                continue

            frames.append(prev_ev)
            added = float(pd.to_numeric(prev_ev["quantity"], errors="coerce").fillna(0.0).sum())
            qty_sum += added
            print(f"[추적] {product}: {year}년 기준 부족/미존재 -> {yy}년 입고 추가 ({added}kg)")

            if min_required_qty is None:
                break
            if qty_sum + 1e-9 >= float(min_required_qty):
                break

    if not frames:
        return pd.DataFrame()

    out = pd.concat(frames, ignore_index=True)
    out = out.sort_values(["event_date"]).reset_index(drop=True)
    return out


# ======================
# 판매 집계
# ======================

def filter_excel_by_ui(result: dict, df: pd.DataFrame | None = None) -> pd.DataFrame:
    file_path = Path(result["file_path"])
    year = int(result["year"])

    if df is None:
        df = load_inventory_df(str(file_path), year)
        if df.empty:
            raise ValueError(f"엑셀에 '{year}' 시트 없음")

    start = pd.Timestamp(f"{year}-{int(result['start_month']):02d}-{int(result['start_day']):02d}")
    end = pd.Timestamp(f"{year}-{int(result['end_month']):02d}-{int(result['end_day']):02d}")

    sdf = df[(df["__event_date__"] >= start) & (df["__event_date__"] <= end)].copy()

    out = (
        sdf[["__product__", "__sales_qty__"]]
        .rename(columns={"__product__": "product", "__sales_qty__": "quantity"})
    )

    out["product"] = out["product"].replace({"nan": None, "None": None, "": None}).str.strip()
    out = out[out["product"].notna()]
    out = out.groupby("product", as_index=False)["quantity"].sum()
    out = out.sort_values("product", key=lambda s: s.str.lower()).reset_index(drop=True)
    return out


# ======================
# 이벤트 추출
# ======================

def extract_inbound_events(
    file_path: Path,
    year: int,
    product: str,
    cutoff: pd.Timestamp,
    df: pd.DataFrame | None = None,
) -> pd.DataFrame:
    if df is None:
        df = load_inventory_df(str(file_path), year)
        if df.empty:
            return pd.DataFrame()

    supplier_s = df["__supplier__"] if "__supplier__" in df.columns else pd.Series("", index=df.index)
    customer_s = df["__customer__"] if "__customer__" in df.columns else pd.Series("", index=df.index)
    note_s = df["__note__"] if "__note__" in df.columns else pd.Series("", index=df.index)

    out = pd.DataFrame({
        "product": df["__product__"],
        "event_date": pd.to_datetime(df["__in_date__"], errors="coerce"),
        "quantity": pd.to_numeric(df["__in_qty__"], errors="coerce").fillna(0.0),
        "unit_price": pd.to_numeric(df["__unit_price__"], errors="coerce"),
        "exchange_rate": pd.to_numeric(df["__exchange_rate__"], errors="coerce"),
        "supplier": supplier_s,
        "customer": customer_s,
        "note": note_s,
    })
    out["is_return"] = (
        out["supplier"].map(_is_return_text)
        | out["customer"].map(_is_return_text)
        | out["note"].map(_is_return_text)
    )

    out = out[
        (out["product"] == product)
        & (out["quantity"] > 0)
        & (out["event_date"].notna())
        & (out["event_date"] <= cutoff)
        & (~out["is_return"])
    ].copy()

    return out.sort_values(["event_date"]).reset_index(drop=True)


def _pick_adjustment_seed(events: pd.DataFrame, prefer_latest: bool) -> pd.Series | None:
    if events.empty:
        return None
    # 보정 lot은 단가/환율이 일부라도 있는 이벤트를 우선 사용.
    # 둘 다 비어있다면 날짜 정확도를 위해 실입고 이벤트 날짜를 seed로 사용한다.
    priced = events[(events["unit_price"].notna()) | (events["exchange_rate"].notna())]
    pool = priced if not priced.empty else events
    if pool.empty:
        return None
    pool = pool.sort_values("event_date")
    return pool.iloc[-1] if prefer_latest else pool.iloc[0]


def extract_sales_events(
    file_path: Path,
    year: int,
    product: str,
    start: pd.Timestamp,
    end: pd.Timestamp,
    df: pd.DataFrame | None = None,
) -> pd.DataFrame:
    if df is None:
        df = load_inventory_df(str(file_path), year)
        if df.empty:
            return pd.DataFrame()

    out = pd.DataFrame({
        "product": df["__product__"],
        "event_date": pd.to_datetime(df["__event_date__"], errors="coerce"),
        "quantity": pd.to_numeric(df["__sales_qty__"], errors="coerce").fillna(0.0),
    })

    out = out[
        (out["product"] == product)
        & (out["quantity"] > 0)
        & (out["event_date"].notna())
        & (out["event_date"] >= start)
        & (out["event_date"] <= end)
    ].copy()

    return out.sort_values(["event_date"]).reset_index(drop=True)


# ======================
# FIFO 코어
# ======================

def consume_fifo_lots(lots: list[dict], qty: float) -> tuple[list[dict], list[dict], float]:
    remain = float(qty)
    allocations = []

    for lot in lots:
        if remain <= 1e-12:
            break

        avail = float(lot["quantity"])
        if avail <= 1e-12:
            continue

        take = min(avail, remain)
        lot["quantity"] = avail - take
        remain -= take

        allocations.append({
            "in_date": lot["in_date"],
            "taken_qty": take,
            "unit_price": lot["unit_price"],
            "exchange_rate": lot["exchange_rate"],
        })

    updated_lots = [x for x in lots if float(x["quantity"]) > 1e-12]
    return updated_lots, allocations, remain


def _only_priced_allocs(allocations: list[dict]) -> list[dict]:
    bucket: dict[tuple[str, float | None, float | None], float] = {}
    for a in allocations:
        in_date = a["in_date"].strftime("%Y-%m-%d") if isinstance(a["in_date"], pd.Timestamp) else str(a["in_date"])
        unit_price = None if a["unit_price"] is None else float(a["unit_price"])
        exchange_rate = None if a["exchange_rate"] is None else float(a["exchange_rate"])
        key = (in_date, unit_price, exchange_rate)
        bucket[key] = bucket.get(key, 0.0) + float(a["taken_qty"])

    out = [
        {
            "in_date": k[0],
            "taken_qty": v,
            "unit_price": k[1],
            "exchange_rate": k[2],
        }
        for k, v in bucket.items()
    ]
    out.sort(key=lambda x: (x["in_date"], x["unit_price"], x["exchange_rate"]))
    return out


def _impute_alloc_pricing_from_inbound(allocations: list[dict], inbound_all: pd.DataFrame) -> list[dict]:
    """단가/환율 누락 allocation에 대해, 같은 품목의 직전 유효 입고값으로 보정."""
    if not allocations:
        return allocations
    if inbound_all.empty:
        return allocations

    src = inbound_all.copy()
    src["event_date"] = pd.to_datetime(src["event_date"], errors="coerce")
    src["unit_price"] = pd.to_numeric(src["unit_price"], errors="coerce")
    src["exchange_rate"] = pd.to_numeric(src["exchange_rate"], errors="coerce")
    src = src.sort_values("event_date").reset_index(drop=True)

    def _latest_before(dt: pd.Timestamp, field: str):
        q = src[(src["event_date"].notna()) & (src["event_date"] <= dt) & (src[field].notna())]
        if q.empty:
            q = src[src[field].notna()]
        if q.empty:
            return None
        return float(q.iloc[-1][field])

    out = []
    for a in allocations:
        in_date = a["in_date"] if isinstance(a["in_date"], pd.Timestamp) else pd.to_datetime(a["in_date"], errors="coerce")
        unit = a.get("unit_price")
        rate = a.get("exchange_rate")

        if unit is None and pd.notna(in_date):
            unit = _latest_before(in_date, "unit_price")
        if rate is None and pd.notna(in_date):
            rate = _latest_before(in_date, "exchange_rate")

        out.append({
            "in_date": a["in_date"],
            "taken_qty": a["taken_qty"],
            "unit_price": unit,
            "exchange_rate": rate,
        })
    return out


def fifo_for_product(
    file_path: Path,
    product: str,
    year: int,
    sales_qty: float,
    start: pd.Timestamp,
    cutoff: pd.Timestamp,
    current_stock: float = 0.0,
    df: pd.DataFrame | None = None,
):
    """
    이벤트 기반 FIFO
    1) 시작 전 입고/출고를 실제 FIFO로 반영하여 시작 시점 잔량 로트 구성
    2) 기간 중 입고/출고를 날짜순으로 반영
    3) 가격 있는 로트만 결과로 반환
    """
    if df is None:
        df = load_inventory_df(str(file_path), year)
        if df.empty:
            return []

    boundary = start - pd.Timedelta(days=1)

    inbound_all = _load_inbound_events_progressive(
        file_path=file_path,
        year=year,
        product=product,
        cutoff=cutoff,
        base_df=df,
        min_required_qty=float(sales_qty),
    )
    if inbound_all.empty:
        print(f"[경고] {product}: cutoff 이전 입고 없음")
        return []

    inbound_pre = inbound_all[inbound_all["event_date"] <= boundary].copy()
    inbound_period = inbound_all[inbound_all["event_date"] >= start].copy()

    sales_pre = extract_sales_events(
        file_path,
        year,
        product,
        start=pd.Timestamp(year=year, month=1, day=1),
        end=boundary,
        df=df,
    )
    sales_period = extract_sales_events(
        file_path,
        year,
        product,
        start=start,
        end=cutoff,
        df=df,
    )

    # 시작 전 로트 큐 구성
    lots: list[dict] = []
    for _, row in inbound_pre.iterrows():
        lots.append({
            "in_date": row["event_date"],
            "quantity": float(row["quantity"]),
            "unit_price": None if pd.isna(row["unit_price"]) else float(row["unit_price"]),
            "exchange_rate": None if pd.isna(row["exchange_rate"]) else float(row["exchange_rate"]),
        })

    # 시작 전 출고 소진
    for _, row in sales_pre.iterrows():
        lots, _, _ = consume_fifo_lots(lots, float(row["quantity"]))

    # 시작 직전 장부재고와 FIFO lot 잔량을 맞추어 보정
    # (입출고 수량만으로 설명되지 않는 재고조정/기초재고 이월 차이를 흡수)
    snapshot_stock = get_stock_snapshot_at_or_before(
        file_path=file_path,
        year=year,
        product=product,
        asof=boundary,
        df=df,
    )
    traced_stock = float(sum(float(x["quantity"]) for x in lots))
    delta = traced_stock - float(snapshot_stock)

    if delta > 1e-9:
        # 추적 잔량이 장부보다 크면 FIFO 순서로 추가 소진
        lots, _, _ = consume_fifo_lots(lots, delta)
    elif delta < -1e-9:
        # 추적 잔량이 장부보다 작으면 보정 lot 추가.
        # 우선순위:
        # 1) 시작 전 가장 이른 입고 seed (FIFO 일관성)
        # 2) 기간 내 가장 이른 입고 seed (시작 전 seed가 전혀 없을 때)
        # 3) 그래도 없으면 boundary
        seed = _pick_adjustment_seed(inbound_pre, prefer_latest=False)
        if seed is None:
            seed = _pick_adjustment_seed(inbound_period, prefer_latest=False)

        if seed is not None:
            adj_in_date = pd.to_datetime(seed["event_date"], errors="coerce")
            adj_unit = None if pd.isna(seed.get("unit_price")) else float(seed["unit_price"])
            adj_rate = None if pd.isna(seed.get("exchange_rate")) else float(seed["exchange_rate"])
        else:
            adj_in_date = boundary
            adj_unit = None
            adj_rate = None

        lots.append({
            "in_date": adj_in_date if pd.notna(adj_in_date) else boundary,
            "quantity": -delta,
            "unit_price": adj_unit,
            "exchange_rate": adj_rate,
        })
        lots.sort(key=lambda x: x["in_date"])

    # 기간 중 이벤트 결합
    period_events = []

    for _, row in inbound_period.iterrows():
        period_events.append({
            "kind": "in",
            "event_date": row["event_date"],
            "quantity": float(row["quantity"]),
            "unit_price": None if pd.isna(row["unit_price"]) else float(row["unit_price"]),
            "exchange_rate": None if pd.isna(row["exchange_rate"]) else float(row["exchange_rate"]),
        })

    for _, row in sales_period.iterrows():
        period_events.append({
            "kind": "out",
            "event_date": row["event_date"],
            "quantity": float(row["quantity"]),
        })

    # 같은 날짜면 입고 먼저, 출고 나중
    period_events.sort(key=lambda x: (x["event_date"], 0 if x["kind"] == "in" else 1))

    period_allocs = []

    for ev in period_events:
        if ev["kind"] == "in":
            lots.append({
                "in_date": ev["event_date"],
                "quantity": ev["quantity"],
                "unit_price": ev["unit_price"],
                "exchange_rate": ev["exchange_rate"],
            })
            lots.sort(key=lambda x: x["in_date"])
        else:
            lots, allocs, remain = consume_fifo_lots(lots, ev["quantity"])
            period_allocs.extend(allocs)
            if remain > 1e-9:
                print(f"[주의] {product}: {ev['event_date'].date()} 출고 {ev['quantity']}kg 중 부족분 {remain}kg")

    imputed_allocs = _impute_alloc_pricing_from_inbound(period_allocs, inbound_all)
    filtered_allocs = _only_priced_allocs(imputed_allocs)

    total = sum(a["taken_qty"] for a in filtered_allocs)
    print(f"[결과] {product} | FIFO(event) | 기간배분건수={len(filtered_allocs)} | 합계={total} | 기대값={sales_qty}")
    return filtered_allocs


# ======================
# 보조 집계/스냅샷
# ======================

def load_inbound_history(file_path: Path, year: int, product: str, df: pd.DataFrame | None = None) -> pd.DataFrame:
    return extract_inbound_events(file_path, year, product, cutoff=pd.Timestamp.max, df=df).rename(columns={"event_date": "in_date"})


def fifo_allocate(sales_qty: float, inbound_df: pd.DataFrame):
    allocations, remaining = [], sales_qty
    for _, row in inbound_df.iterrows():
        if remaining <= 0:
            break
        take = min(remaining, row["quantity"])
        allocations.append({
            "in_date": row["in_date"],
            "taken_qty": take,
            "unit_price": row.get("unit_price"),
            "exchange_rate": row.get("exchange_rate"),
        })
        remaining -= take
    return allocations, remaining

# 분석용 scripts > check_ambn.py
def sum_sales_in_range(
    file_path: Path,
    year: int,
    product: str,
    start: pd.Timestamp,
    end: pd.Timestamp,
    df: pd.DataFrame | None = None,
) -> float:
    ev = extract_sales_events(file_path, year, product, start, end, df=df)
    return float(pd.to_numeric(ev["quantity"], errors="coerce").fillna(0.0).sum()) if not ev.empty else 0.0


def apply_fifo_consumption(lots_df: pd.DataFrame, consume_qty: float):
    allocations = []
    remaining = float(consume_qty)
    lots = lots_df.copy()

    for i in range(len(lots)):
        if remaining <= 0:
            break
        idx = lots.index[i]
        avail = float(lots.at[idx, "quantity"])
        take = min(avail, remaining)
        if take > 0:
            lots.at[idx, "quantity"] = avail - take
            allocations.append({
                "in_date": lots.at[idx, "in_date"],
                "taken_qty": take,
                "unit_price": float(lots.at[idx, "unit_price"]) if lots.at[idx, "unit_price"] is not None else None,
                "exchange_rate": float(lots.at[idx, "exchange_rate"]) if lots.at[idx, "exchange_rate"] is not None else None,
            })
            remaining -= take

    return lots, allocations, remaining


def get_current_stock(
    file_path: Path,
    year: int,
    product: str,
    start: pd.Timestamp,
    end: pd.Timestamp,
    df: pd.DataFrame | None = None,
) -> float:
    try:
        if df is None:
            df = load_inventory_df(str(file_path), year)
            if df.empty:
                return 0.0

        sdf = df[(df["__product__"] == product) & (df["__row_date__"] >= start) & (df["__row_date__"] <= end)].copy()
        if sdf.empty:
            return 0.0
        last_row = sdf.sort_values("__row_date__").iloc[-1]
        return 0.0 if pd.isna(last_row["__stock__"]) else float(last_row["__stock__"])
    except Exception as e:
        print(f"[경고] {product} 현재고 추출 실패: {e}")
        return 0.0


def get_stock_snapshot_at_or_before(
    file_path: Path,
    year: int,
    product: str,
    asof: pd.Timestamp,
    df: pd.DataFrame | None = None,
) -> float:
    file_path_str = str(file_path)
    if df is None:
        df = load_inventory_df(file_path_str, year)
        if df.empty:
            return 0.0

    sdf = df[(df["__product__"] == product) & (df["__row_date__"] <= asof)].copy()
    if sdf.empty:
        # 요청 연도에 이력이 없으면 과거 연도로 역추적
        for yy in _years_to_scan_desc(file_path_str, int(year)):
            if yy >= int(year):
                continue
            d_prev = load_inventory_df(file_path_str, int(yy))
            if d_prev.empty:
                continue
            s_prev = d_prev[(d_prev["__product__"] == product) & (d_prev["__row_date__"] <= asof)].copy()
            if s_prev.empty:
                continue
            last_prev = s_prev.sort_values("__row_date__").iloc[-1]
            print(f"[추적] {product}: 재고스냅샷을 {yy}년 시트에서 보강")
            return 0.0 if pd.isna(last_prev["__stock__"]) else float(last_prev["__stock__"])
        return 0.0
    last_row = sdf.sort_values("__row_date__").iloc[-1]
    return 0.0 if pd.isna(last_row["__stock__"]) else float(last_row["__stock__"])


def sum_inbound_until(
    file_path: Path,
    year: int,
    product: str,
    until: pd.Timestamp,
    df: pd.DataFrame | None = None,
) -> float:
    if df is None:
        df = load_inventory_df(str(file_path), year)
        if df.empty:
            return 0.0

    mask = (df["__product__"] == product) & (df["__in_date__"] <= until)
    return float(pd.to_numeric(df.loc[mask, "__in_qty__"], errors="coerce").fillna(0.0).sum())


# ======================
# 결과 요약
# ======================

def summarize_allocations(allocations: list[dict]) -> pd.DataFrame:
    if not allocations:
        return pd.DataFrame(columns=["unit_price", "exchange_rate", "in_date", "taken_qty"])

    df = pd.DataFrame(allocations)
    out = (
        df.groupby(["unit_price", "exchange_rate", "in_date"], as_index=False)["taken_qty"]
        .sum()
        .sort_values(["in_date", "unit_price", "exchange_rate"])
        .reset_index(drop=True)
    )
    return out


# ======================
# 최종 실행
# ======================

def filter_and_fifo(result: dict):
    file_path = Path(result["file_path"])
    year = int(result["year"])

    start = pd.Timestamp(f"{year}-{int(result['start_month']):02d}-{int(result['start_day']):02d}")
    end = pd.Timestamp(f"{year}-{int(result['end_month']):02d}-{int(result['end_day']):02d}")

    df_all = load_inventory_df(str(file_path), year)
    sales_df = filter_excel_by_ui(result, df=df_all)
    final_results = []

    for _, row in sales_df.iterrows():
        product = row["product"]
        qty = float(row["quantity"])

        current_stock = get_current_stock(file_path, year, product, start, end, df=df_all)

        print(f"DEBUG - FIFO 처리중: {product}, 판매수량={qty}, 현재고(보고)={current_stock}")
        allocations = fifo_for_product(
            file_path,
            product,
            year,
            qty,
            start=start,
            cutoff=end,
            current_stock=current_stock,
            df=df_all,
        )

        final_results.append({
            "product": product,
            "sales_qty": qty,
            "current_stock": current_stock,
            "total_consumption": qty + current_stock,
            "fifo_allocations": allocations,
        })

    return final_results
