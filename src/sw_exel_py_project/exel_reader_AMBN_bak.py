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

    date_col = _pick_date_col(df, INVENTORY_COLUMN_MAP["date"])
    prod_col = _pick_product_col(df, INVENTORY_COLUMN_MAP["product_name"])

    df["product"] = df.loc[:, prod_col].astype(str).str.strip()
    df[date_col] = df[date_col].apply(lambda x: _parse_korean_md(x, year))

    start = pd.Timestamp(f"{year}-{int(result['start_month']):02d}-{int(result['start_day']):02d}")
    end   = pd.Timestamp(f"{year}-{int(result['end_month']):02d}-{int(result['end_day']):02d}")
    df = df[(df[date_col] >= start) & (df[date_col] <= end)]

    qty_cols = _pick_sales_qty_cols(df)
    df["__qty__"] = (df[qty_cols]
                     .replace(r"[^\d\.\-]", "", regex=True)
                     .apply(pd.to_numeric, errors="coerce")
                     .fillna(0.0).sum(axis=1))

    out = (df[["product", "__qty__"]]
           .rename(columns={"__qty__": "quantity"}))

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

    # 입고 수량
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

    # 단가/환율
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

    # 원가 확정된 입고만, 해당 제품만
    inbound = (inbound.query("product == @product and quantity > 0")
                      .dropna(subset=["in_date"])
                      .copy())
    inbound = inbound[(inbound["unit_price"].notna()) & (inbound["exchange_rate"].notna())]

    # ✅ FIFO: 오래된 것부터 (ascending=True)
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

def sum_sales_in_range(file_path: Path, year: int, product: str, start: pd.Timestamp, end: pd.Timestamp) -> float:
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

    df[date_col] = df[date_col].apply(lambda x: _parse_korean_md(x, year))
    df["product"] = df[prod_col].astype(str).str.strip()

    # 납품/매출 수량 컬럼 찾기
    qty_cols = _pick_sales_qty_cols(df)
    df["__qty__"] = (df[qty_cols]
                     .replace(r"[^\d\.\-]", "", regex=True)
                     .apply(pd.to_numeric, errors="coerce")
                     .fillna(0.0).sum(axis=1))

    mask = (df["product"] == product) & (df[date_col] >= start) & (df[date_col] <= end)
    return float(df.loc[mask, "__qty__"].sum())

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
    start: pd.Timestamp,       # 기간 시작 (예: 2025-07-24)
    cutoff: pd.Timestamp,      # 기간 종료 (예: 2025-08-25)
    current_stock: float = 0.0 # 보고용 (배분엔 사용 안 함)
):
    # 0) cutoff 이전 모든 입고 로트를 오래된→최신(asc)으로
    inbound_all = load_inbound_history(file_path, year, product)
    inbound_all = inbound_all[inbound_all["in_date"] <= cutoff].sort_values("in_date", ascending=True).reset_index(drop=True)
    if inbound_all.empty:
        print(f"[경고] {product}: cutoff 이전 입고 없음")
        return []

    # 1) 시작일 스냅샷 현재고
    start_snapshot = get_stock_snapshot_at_or_before(file_path, year, product, start - pd.Timedelta(days=1))

    # 2) 시작일 이전 총 입고량
    inbound_before_start = sum_inbound_until(file_path, year, product, start - pd.Timedelta(days=1))

    # 3) 시작 전 소진량 = 이전 총입고 - 시작 스냅샷현재고 (음수면 0)
    pre_consumption = max(0.0, inbound_before_start - start_snapshot)

    # 4) 시작 전 소진을 FIFO로 차감 → '기간 시작 시점' 로트 잔량 형성
    lots_after_pre, _, pre_remain = apply_fifo_consumption(inbound_all.copy(), pre_consumption)
    if pre_remain > 1e-9:
        print(f"[주의] {product}: 시작 전 소진({pre_consumption}) 차감 중 입고 부족(잔여 {pre_remain})")

    # 5) 기간 내 판매량을 FIFO로 배분
    lots_after_period, period_allocs, remain = apply_fifo_consumption(lots_after_pre, float(sales_qty))
    if remain > 1e-9:
        print(f"[주의] {product}: 기간 판매({sales_qty}) 배분 중 입고 부족(잔여 {remain})")

    # 6) 날짜 문자열화
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