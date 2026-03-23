import re
from pathlib import Path
import pandas as pd
import openpyxl
import unicodedata
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
        except: 
            return pd.NaT
    return pd.to_datetime(txt, errors="coerce")

def _norm(s): 
    return "" if pd.isna(s) else str(s).strip().lower()

def _is_unnamed(s): 
    return _norm(s).startswith("unnamed")


# ======================
# 컬럼 선택기
# ======================

# 날짜
def _pick_date_col(df, col_date_lv0):
    for c in df.columns:
        if isinstance(c, tuple) and _norm(c[0]) == _norm(col_date_lv0): 
            return c
        if not isinstance(c, tuple) and _norm(c) == _norm(col_date_lv0): 
            return c
    raise KeyError(f"날짜 컬럼(level=0='{col_date_lv0}') 없음")

def _pick_arrival_col(df, col_arrival_lv0: str):
    """도착일(보조 날짜) 컬럼 선택. 없으면 None"""
    for c in df.columns:
        if isinstance(c, tuple) and _norm(c[0]) == _norm(col_arrival_lv0):
            return c
        if not isinstance(c, tuple) and _norm(c) == _norm(col_arrival_lv0):
            return c
    return None

# 품명
def _pick_product_col(df, col_prod_lv0):
    for c in df.columns:
        if isinstance(c, tuple) and _norm(c[0]) == _norm(col_prod_lv0): 
            return c
        if not isinstance(c, tuple) and _norm(c) == _norm(col_prod_lv0): 
            return c
    raise KeyError(f"품명 컬럼(level=0='{col_prod_lv0}') 없음")

# 마지막 재고
def _pick_current_stock_col(df, col_stock_lv0):
    for c in df.columns:
        if isinstance(c, tuple) and _norm(c[0]) == _norm(col_stock_lv0): 
            return c
        if not isinstance(c, tuple) and _norm(c) == _norm(col_stock_lv0): 
            return c
    raise KeyError(f"현재고 컬럼(level=0='{col_stock_lv0}') 없음")

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

def _pick_inbound_unit_price_cols(df: pd.DataFrame):
    """입고/수입 단가 컬럼"""
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

def _make_event_date_series(df: pd.DataFrame, primary_col, arrival_col, year: int) -> pd.Series:
    """기본: '날 짜' → NaT인 행만 '도착일'로 폴백"""
    s1 = df[primary_col].apply(lambda x: _parse_korean_md(x, year))
    if arrival_col is None:
        return s1
    s2 = df[arrival_col].apply(lambda x: _parse_korean_md(x, year))
    return s1.where(~s1.isna(), s2)

def _pick_inbound_exchange_rate_cols(df: pd.DataFrame):
    """입고/수입 환율 컬럼"""
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
# 판매 집계
# ======================

def filter_excel_by_ui(result: dict) -> pd.DataFrame:
    file_path = Path(result["file_path"])
    year = int(result["year"])

    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    sheets = [s for s in wb.sheetnames if str(year) in s]
    if not sheets:
        raise ValueError(f"엑셀에 '{year}' 시트 없음")

    df = pd.read_excel(
        file_path, sheet_name=sheets[0],
        engine="openpyxl", header=[2, 3], dtype=str
    )

    date_col    = _pick_date_col(df, INVENTORY_COLUMN_MAP["date"])
    prod_col    = _pick_product_col(df, INVENTORY_COLUMN_MAP["product_name"])
    arrival_col = _pick_arrival_col(df, INVENTORY_COLUMN_MAP.get("arrival_date", "도착일"))

    df["__product__"]    = df.loc[:, prod_col].astype(str).str.strip()
    df["__event_date__"] = _make_event_date_series(df, date_col, arrival_col, year)

    start = pd.Timestamp(f"{year}-{int(result['start_month']):02d}-{int(result['start_day']):02d}")
    end   = pd.Timestamp(f"{year}-{int(result['end_month']):02d}-{int(result['end_day']):02d}")
    df = df[(df["__event_date__"] >= start) & (df["__event_date__"] <= end)]

    qty_cols = _pick_sales_qty_cols(df)
    df["__qty__"] = (df[qty_cols]
                     .replace(r"[^\d\.\-]", "", regex=True)
                     .apply(pd.to_numeric, errors="coerce")
                     .fillna(0.0).sum(axis=1))

    out = (df[["__product__", "__qty__"]]
           .rename(columns={"__product__": "product", "__qty__": "quantity"}))

    out["product"] = out["product"].replace({"nan": None, "None": None, "": None}).str.strip()
    out = out[out["product"].notna()]

    out = out.groupby("product", as_index=False)["quantity"].sum()
    out = out.sort_values("product", key=lambda s: s.str.lower()).reset_index(drop=True)
    return out


# ======================
# FIFO 관련
# ======================

def load_inbound_history(file_path: Path, year: int, product: str) -> pd.DataFrame:
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    sheets = [s for s in wb.sheetnames if str(year) in s]
    if not sheets:
        return pd.DataFrame()

    df = pd.read_excel(file_path, sheet_name=sheets[0], engine="openpyxl", header=[2, 3], dtype=str)
    date_col = _pick_date_col(df, INVENTORY_COLUMN_MAP["date"])
    prod_col = _pick_product_col(df, INVENTORY_COLUMN_MAP["product_name"])
    df[date_col] = df[date_col].apply(lambda x: _parse_korean_md(x, year))

    # 입고 수량 컬럼
    qty_cols = []
    for c in df.columns:
        if isinstance(c, tuple):
            parts = [_canon(x) for x in c]
            if ("입고" in parts[0] or "수입" in parts[0]) and any(("수량" in p) or ("kg" in p) for p in parts[1:]):
                qty_cols.append(c)
    if not qty_cols:
        return pd.DataFrame()

    df["__in_qty__"] = (df[qty_cols].replace(r"[^\d\.\-]", "", regex=True)
                                     .apply(pd.to_numeric, errors="coerce")
                                     .fillna(0.0).sum(axis=1))

    # 단가/환율 (없어도 수량은 유지)
    price_cols = _pick_inbound_unit_price_cols(df)
    rate_cols  = _pick_inbound_exchange_rate_cols(df)
    unit_price = (df[price_cols].replace(r"[^\d\.\-]", "", regex=True)
                                .apply(pd.to_numeric, errors="coerce")
                                .mean(axis=1) if price_cols else None)
    exchange   = (df[rate_cols].replace(r"[^\d\.\-]", "", regex=True)
                               .apply(pd.to_numeric, errors="coerce")
                               .mean(axis=1) if rate_cols else None)

    inbound = pd.DataFrame({
        "product": df[prod_col].astype(str).str.strip(),
        "in_date": df[date_col],
        "quantity": df["__in_qty__"],
        "unit_price": unit_price if unit_price is not None else None,
        "exchange_rate": exchange if exchange is not None else None,
    })

    inbound = (inbound.query("product == @product and quantity > 0")
                      .dropna(subset=["in_date"])
                      .copy())

    # ✅ 환율/단가 NaN이어도 드롭하지 않음 (원가 집계 시에만 주의)
    return inbound.sort_values("in_date", ascending=True).reset_index(drop=True)

def fifo_allocate(sales_qty: float, inbound_df: pd.DataFrame):
    allocations, remaining = [], sales_qty
    for _, row in inbound_df.iterrows():
        if remaining <= 0: break
        take = min(remaining, row["quantity"])
        allocations.append({
            "in_date": row["in_date"],
            "taken_qty": take,
            "unit_price": row.get("unit_price"),
            "exchange_rate": row.get("exchange_rate")
        })
        remaining -= take
    return allocations, remaining

def sum_sales_in_range(file_path: Path, year: int, product: str,
                       start: pd.Timestamp, end: pd.Timestamp) -> float:
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    sheets = [s for s in wb.sheetnames if str(year) in s]
    if not sheets:
        return 0.0

    df = pd.read_excel(
        file_path, sheet_name=sheets[0],
        engine="openpyxl", header=[2, 3], dtype=str
    )

    date_col    = _pick_date_col(df, INVENTORY_COLUMN_MAP["date"])
    arrival_col = _pick_arrival_col(df, INVENTORY_COLUMN_MAP.get("arrival_date", "도착일"))
    prod_col    = _pick_product_col(df, INVENTORY_COLUMN_MAP["product_name"])

    # 내부 안전 컬럼으로 복사
    df["__product__"]    = df[prod_col].astype(str).str.strip()
    df["__event_date__"] = _make_event_date_series(df, date_col, arrival_col, year)

    qty_cols = _pick_sales_qty_cols(df)
    df["__qty__"] = (df[qty_cols].replace(r"[^\d\.\-]", "", regex=True)
                                .apply(pd.to_numeric, errors="coerce")
                                .fillna(0.0).sum(axis=1))

    mask = (df["__product__"] == product) & (df["__event_date__"] >= start) & (df["__event_date__"] <= end)
    total = float(df.loc[mask, "__qty__"].sum())

    # 디버그(원인 파악용)
    # print(f"[DBG] sum_sales_in_range {product} {start.date()}~{end.date()} = {total}")
    return total

def apply_fifo_consumption(lots_df: pd.DataFrame, consume_qty: float):
    """
    lots_df: in_date asc, columns = [in_date, quantity, unit_price, exchange_rate]
    consume_qty: 차감할 출고 합계
    return (lots_df_after, allocations_list)
    """
    allocations = []
    remaining = float(consume_qty)

    # 내부 복사본
    lots = lots_df.copy()

    for i in range(len(lots)):
        if remaining <= 0:
            break
        avail = float(lots.at[i, "quantity"])
        take = min(avail, remaining)
        if take > 0:
            lots.at[i, "quantity"] = avail - take
            allocations.append({
                "in_date": lots.at[i, "in_date"],
                "taken_qty": take,
                "unit_price": float(lots.at[i, "unit_price"]) if lots.at[i, "unit_price"] is not None else None,
                "exchange_rate": float(lots.at[i, "exchange_rate"]) if lots.at[i, "exchange_rate"] is not None else None,
            })
            remaining -= take

    return lots, allocations, remaining

def _pick_current_stock_col(df: pd.DataFrame, col_stock_name: str):
    """현재고 컬럼 찾기"""
    for c in df.columns:
        if isinstance(c, tuple):
            # 멀티헤더에서 어느 레벨에든 현재고 키워드가 있으면
            for level_text in c:
                if _norm(level_text) == _norm(col_stock_name):
                    return c
        else:
            if _norm(c) == _norm(col_stock_name):
                return c
    raise KeyError(f"현재고 컬럼('{col_stock_name}') 없음")

def get_current_stock(file_path: Path, year: int, product: str, start: pd.Timestamp, end: pd.Timestamp) -> float:
    """해당 제품의 기간 내 마지막 현재고 추출"""
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        sheets = [s for s in wb.sheetnames if str(year) in s]
        if not sheets:
            return 0.0
        df = pd.read_excel(
            file_path, sheet_name=sheets[0],
            engine="openpyxl", header=[2, 3], dtype=str
        )

        date_col = _pick_date_col(df, INVENTORY_COLUMN_MAP["date"])
        prod_col = _pick_product_col(df, INVENTORY_COLUMN_MAP["product_name"])
        stock_col = _pick_current_stock_col(df, INVENTORY_COLUMN_MAP["current_stock"])

        df[date_col] = df[date_col].apply(lambda x: _parse_korean_md(x, year))
        df["product"] = df[prod_col].astype(str).str.strip()

         # 해당 제품 + 기간 내 데이터
        mask = (df["product"] == product) & (df[date_col] >= start) & (df[date_col] <= end)
        product_df = df[mask]
        
        if product_df.empty:
            return 0.0
        
        # 마지막 행의 현재고
        last_row = product_df.sort_values(date_col).iloc[-1]
        stock_value = pd.to_numeric(last_row[stock_col], errors='coerce')
        
        return 0.0 if pd.isna(stock_value) else float(stock_value)
        
    except Exception as e:
        print(f"[경고] {product} 현재고 추출 실패: {e}")
        return 0.0

def fifo_for_product(
    file_path: Path,
    product: str,
    year: int,
    sales_qty: float,          # 기간 내 판매량 (예: 7/24~8/25 합계)
    start: pd.Timestamp,       # 기간 시작
    cutoff: pd.Timestamp,      # 기간 종료
    current_stock: float = 0.0 # 보고용
):
    """
    규칙:
      - 선소진(pre-consumption)은 시작일 직전 로트(<= start-1)에서만 FIFO 차감
      - 선소진 규모는 하이브리드: max(실제판매합계, 입고합계-시작스냅샷현재고)
      - 부족 시 오프닝(가상) 로트로 흡수 (기간 배분에는 제외)
      - 기간 배분 풀 = (선소진 후 '시작 전' 잔량) + ('시작 이후' 로트)
    """

    # 0) cutoff 이전 모든 입고 로트(오래된→최신)
    inbound_all = load_inbound_history(file_path, year, product)
    inbound_all = (
        inbound_all[inbound_all["in_date"] <= cutoff]
        .sort_values("in_date", ascending=True)
        .reset_index(drop=True)
    )
    if inbound_all.empty:
        print(f"[경고] {product}: cutoff 이전 입고 없음")
        return []

    # 경계일
    boundary = start - pd.Timedelta(days=1)

    # 1) 선소진 규모 계산 (하이브리드)
    pre_sales_sum = sum_sales_in_range(
        file_path, year, product,
        pd.Timestamp(year=year, month=1, day=1),
        boundary
    )
    start_snapshot = get_stock_snapshot_at_or_before(file_path, year, product, boundary)
    inbound_before_start = sum_inbound_until(file_path, year, product, boundary)

    # 역산값 (음수 방지)
    pre_backcalc = max(0.0, float(inbound_before_start) - float(start_snapshot))

    # 하이브리드: 둘 중 큰 값 채택
    pre_consumption = max(float(pre_sales_sum), float(pre_backcalc))

    # 2) 시작 전/후 로트 분리
    pre_lots  = inbound_all[inbound_all["in_date"] <= boundary].copy()
    post_lots = inbound_all[inbound_all["in_date"] >  boundary].copy()

    # 3) 오프닝(기초) 로트 생성(필요시)
    pre_capacity = float(pre_lots["quantity"].sum()) if not pre_lots.empty else 0.0
    opening_need = max(0.0, pre_consumption - pre_capacity)
    if opening_need > 0:
        opening_lot = pd.DataFrame([{
            "product": product,
            "in_date": start - pd.Timedelta(days=100000),  # 매우 과거
            "quantity": opening_need,
            "unit_price": None,
            "exchange_rate": None,
            "_is_opening": True,
        }])
        pre_lots = pre_lots.copy()
        pre_lots["_is_opening"] = False
        pre_lots = pd.concat([opening_lot, pre_lots], ignore_index=True)
    else:
        pre_lots = pre_lots.copy()
        pre_lots["_is_opening"] = False

    # 4) 선소진 적용 (오직 pre_lots에서만)
    pre_lots_sorted = pre_lots.sort_values("in_date", ascending=True).reset_index(drop=True)
    lots_after_pre, _, pre_remaining = apply_fifo_consumption(pre_lots_sorted, float(pre_consumption))
    if pre_remaining > 1e-9:
        print(f"[주의] {product}: 시작 전 소진({pre_consumption}) 차감 중 입고 부족(잔여 {pre_remaining})")

    # 0/미소량 제거
    lots_after_pre = lots_after_pre[pd.to_numeric(lots_after_pre["quantity"], errors="coerce").fillna(0.0) > 1e-12].copy()

    # 오프닝 로트는 기간 배분에서 제외
    if "_is_opening" in lots_after_pre.columns:
        lots_after_pre = lots_after_pre[~lots_after_pre["_is_opening"]].copy()
        lots_after_pre.drop(columns=["_is_opening"], inplace=True, errors="ignore")

    # 5) 기간 배분 풀 = (선소진 후 pre 잔량) + (post_lots)
    period_pool = pd.concat([lots_after_pre, post_lots], ignore_index=True)
    if not period_pool.empty:
        period_pool = (
            period_pool.sort_values("in_date", ascending=True)
                       .reset_index(drop=True)
        )
        period_pool["quantity"] = pd.to_numeric(period_pool["quantity"], errors="coerce").fillna(0.0)
        period_pool = period_pool[period_pool["quantity"] > 1e-12]

    # 디버그(원인 추적)
    # print("[DBG] pre_sales_sum:", pre_sales_sum, "start_snapshot:", start_snapshot, "inbound_before_start:", inbound_before_start, "pre_backcalc:", pre_backcalc, "pre_consumption:", pre_consumption)
    # print("[DBG] pre_lots_head:", pre_lots[["in_date","quantity","unit_price","exchange_rate"]].head(5).to_dict("records"))
    # print("[DBG] post_lots_head:", post_lots[["in_date","quantity","unit_price","exchange_rate"]].head(5).to_dict("records"))
    # print("[DBG] period_pool_head:", period_pool[["in_date","quantity","unit_price","exchange_rate"]].head(8).to_dict("records"))

    # 6) 기간 내 판매량 FIFO 배분
    _, period_allocs, remain = apply_fifo_consumption(period_pool, float(sales_qty))
    if remain > 1e-9:
        print(f"[주의] {product}: 기간 판매({sales_qty}) 배분 중 입고 부족(잔여 {remain})")

    # 7) 날짜 문자열화
    for a in period_allocs:
        if isinstance(a["in_date"], pd.Timestamp):
            a["in_date"] = a["in_date"].strftime("%Y-%m-%d")

    total = sum(a["taken_qty"] for a in period_allocs)
    print(f"[결과] {product} | FIFO | 기간배분건수={len(period_allocs)} | 합계={total} | 기대값={sales_qty}")
    return period_allocs

def get_stock_snapshot_at_or_before(file_path: Path, year: int, product: str, asof: pd.Timestamp) -> float:
    """as-of 날짜(포함) 이전의 마지막 '현재고' 값을 스냅샷으로 얻는다."""
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    sheets = [s for s in wb.sheetnames if str(year) in s]
    if not sheets:
        return 0.0
    df = pd.read_excel(file_path, sheet_name=sheets[0], engine="openpyxl", header=[2, 3], dtype=str)

    date_col  = _pick_date_col(df, INVENTORY_COLUMN_MAP["date"])
    prod_col  = _pick_product_col(df, INVENTORY_COLUMN_MAP["product_name"])
    stock_col = _pick_current_stock_col(df, INVENTORY_COLUMN_MAP["current_stock"])

    df[date_col] = df[date_col].apply(lambda x: _parse_korean_md(x, year))
    df["product"] = df[prod_col].astype(str).str.strip()

    mask = (df["product"] == product) & (df[date_col] <= asof)
    sdf  = df[mask]
    if sdf.empty:
        return 0.0

    last_row = sdf.sort_values(date_col).iloc[-1]
    val = pd.to_numeric(last_row[stock_col], errors="coerce")
    return 0.0 if pd.isna(val) else float(val)

def sum_inbound_until(file_path: Path, year: int, product: str, until: pd.Timestamp) -> float:
    """until(포함) 이전의 해당 품목 총 입고수량(kg)."""
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    sheets = [s for s in wb.sheetnames if str(year) in s]
    if not sheets:
        return 0.0

    df = pd.read_excel(file_path, sheet_name=sheets[0], engine="openpyxl", header=[2, 3], dtype=str)
    date_col = _pick_date_col(df, INVENTORY_COLUMN_MAP["date"])
    prod_col = _pick_product_col(df, INVENTORY_COLUMN_MAP["product_name"])
    df[date_col] = df[date_col].apply(lambda x: _parse_korean_md(x, year))
    df["product"] = df[prod_col].astype(str).str.strip()

    # 입고 수량 컬럼 모으기
    qty_cols = []
    for c in df.columns:
        if isinstance(c, tuple):
            parts = [_canon(x) for x in c]
            if ("입고" in parts[0] or "수입" in parts[0]) and any(("수량" in p) or ("kg" in p) for p in parts[1:]):
                qty_cols.append(c)

    if not qty_cols:
        return 0.0

    df["__in_qty__"] = (df[qty_cols].replace(r"[^\d\.\-]", "", regex=True)
                                    .apply(pd.to_numeric, errors="coerce")
                                    .fillna(0.0).sum(axis=1))
    mask2 = (df["product"] == product) & (df[date_col] <= until)
    return float(df.loc[mask2, "__in_qty__"].sum())
# ======================
# 최종 실행
# ======================

def filter_and_fifo(result: dict):
    file_path = Path(result["file_path"])
    year = int(result["year"])

    start = pd.Timestamp(f"{year}-{int(result['start_month']):02d}-{int(result['start_day']):02d}")
    end   = pd.Timestamp(f"{year}-{int(result['end_month']):02d}-{int(result['end_day']):02d}")

    sales_df = filter_excel_by_ui(result)
    final_results = []

    for _, row in sales_df.iterrows():
        product = row["product"]
        qty = float(row["quantity"])

        current_stock = get_current_stock(file_path, year, product, start, end)  # 보고용

        print(f"DEBUG - FIFO 처리중: {product}, 판매수량={qty}, 현재고(보고)={current_stock}")
        allocations = fifo_for_product(file_path, product, year, qty, start=start, cutoff=end, current_stock=current_stock)
        final_results.append({
            "product": product,
            "sales_qty": qty,
            "current_stock": current_stock,            # 단순 참고값
            "total_consumption": qty + current_stock,  # 단순 참고값 (배분미사용)
            "fifo_allocations": allocations
        })
    return final_results


# DEBUG - FIFO 처리중: AMBN, 판매수량=31200.0
# [처리중] AMBN | 모드=LIFO(equal-lot) | B=16000 | 판매수량=31200.0 | cutoff=2025-08-25
# [결과] AMBN | 모드=LIFO(equal-lot) | 할당건수=3 | 합계=31200.0 | 기대값=31200.0 | 부족분=0.0
# DEBUG - FIFO 결과: [{'in_date': '2025-08-14', 'taken_qty': 16000.0, 'unit_price': 10.85, 'exchange_rate': 1388.1}, {'in_date': '2025-07-29', 'taken_qty': 14400.0, 'unit_price': 10.85, 'exchange_rate': 1372.0}, {'in_date': '2025-07-21', 'taken_qty': 800.0, 'unit_price': 10.85, 'exchange_rate': 1391.8}]