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

def _pick_date_col(df, col_date_lv0):
    for c in df.columns:
        if isinstance(c, tuple) and _norm(c[0]) == _norm(col_date_lv0): 
            return c
        if not isinstance(c, tuple) and _norm(c) == _norm(col_date_lv0): 
            return c
    raise KeyError(f"날짜 컬럼(level=0='{col_date_lv0}') 없음")

def _pick_product_col(df, col_prod_lv0):
    for c in df.columns:
        if isinstance(c, tuple) and _norm(c[0]) == _norm(col_prod_lv0): 
            return c
        if not isinstance(c, tuple) and _norm(c) == _norm(col_prod_lv0): 
            return c
    raise KeyError(f"품명 컬럼(level=0='{col_prod_lv0}') 없음")

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

    # 입고 수량 집계
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

    # ✅ 핵심: 원가 확정(단가, 환율 둘 다 숫자)된 입고만 사용 + 최신→과거 정렬
    inbound = (inbound.query("product == @product and quantity > 0")
                      .dropna(subset=["in_date"])
                      .copy())
    inbound = inbound[(inbound["unit_price"].notna()) & (inbound["exchange_rate"].notna())]

    return inbound.sort_values("in_date", ascending=False).reset_index(drop=True)

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

def fifo_for_product(
    file_path: Path,
    product: str,
    year: int,
    sales_qty: float,
    start: pd.Timestamp,   # 사용은 안 하지만 시그니처 유지
    cutoff: pd.Timestamp,  # 마감일(이후 입고는 무시)
):
    # 1) cutoff 연도까지 시트 모으기
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    year_sheets = []
    for s in wb.sheetnames:
        m = re.search(r"\d{4}", s or "")
        if m:
            yy = int(m.group())
            if yy <= cutoff.year:
                year_sheets.append((yy, s))
    # 최신→과거 순서가 필요하니 나중에 합친 뒤 정렬
    year_sheets.sort()

    # 2) 해당 품목의 '원가 확정된' 입고만 모아서 최신→과거 정렬
    inbound_all = []
    for yy, _ in year_sheets:
        ib = load_inbound_history(file_path, yy, product)  # 이미 원가확정 & 날짜 정렬(내림차순) 되도록 앞서 정의되어 있다고 가정
        if not ib.empty:
            inbound_all.append(ib)
    if not inbound_all:
        print(f"[결과] {product} | 모드=LIFO(equal-lot) | 할당건수=0 | 합계=0 | 기대값={sales_qty} | 부족분={sales_qty}")
        return []

    inbound = pd.concat(inbound_all, ignore_index=True)
    # 마감일 이후 입고는 제외
    inbound = inbound[inbound["in_date"] <= cutoff]
    # 최신→과거
    inbound = inbound.sort_values("in_date", ascending=False).reset_index(drop=True)

    if inbound.empty:
        print(f"[결과] {product} | 모드=LIFO(equal-lot) | 할당건수=0 | 합계=0 | 기대값={sales_qty} | 부족분={sales_qty}")
        return []

    # 3) 배치 크기 B 추정 (최신 입고수량 또는 최빈값)
    try:
        if not inbound["quantity"].mode().empty:
            B = float(inbound["quantity"].mode().iloc[0])
        else:
            B = float(inbound.iloc[0]["quantity"])
    except Exception:
        B = float(inbound.iloc[0]["quantity"])

    remaining = float(sales_qty)
    allocations = []
    idx = 0

    print(f"[처리중] {product} | 모드=LIFO(equal-lot) | B={B:.0f} | 판매수량={remaining} | cutoff={cutoff.date()}")

    # 4) 통 배치 소진 (B 단위로)
    while remaining >= B and idx < len(inbound):
        row = inbound.iloc[idx]
        take = min(B, float(row["quantity"]))  # 혹시 배치 크기와 실제 수량이 다르면 실제 수량 한도 내
        allocations.append({
            "in_date": row["in_date"],
            "taken_qty": take,
            "unit_price": (None if pd.isna(row.get("unit_price")) else float(row.get("unit_price"))),
            "exchange_rate": (None if pd.isna(row.get("exchange_rate")) else float(row.get("exchange_rate"))),
        })
        remaining -= take
        idx += 1

    # 5) 잔량 r 분해
    if remaining > 0 and idx < len(inbound):
        r = remaining  # 0 < r < B (혹은 B보다 작게 남아야 정상)
        row_next = inbound.iloc[idx]

        if r > B / 2 and (idx + 1) < len(inbound):
            # 두 배치 분할: (2r - B) + (B - r) = r
            part2 = max(0.0, min(2 * r - B, float(row_next["quantity"])))
            part3 = max(0.0, min(B - r, float(inbound.iloc[idx + 1]["quantity"])))

            if part2 > 0:
                allocations.append({
                    "in_date": row_next["in_date"],
                    "taken_qty": part2,
                    "unit_price": (None if pd.isna(row_next.get("unit_price")) else float(row_next.get("unit_price"))),
                    "exchange_rate": (None if pd.isna(row_next.get("exchange_rate")) else float(row_next.get("exchange_rate"))),
                })
            if part3 > 0:
                row_next2 = inbound.iloc[idx + 1]
                allocations.append({
                    "in_date": row_next2["in_date"],
                    "taken_qty": part3,
                    "unit_price": (None if pd.isna(row_next2.get("unit_price")) else float(row_next2.get("unit_price"))),
                    "exchange_rate": (None if pd.isna(row_next2.get("exchange_rate")) else float(row_next2.get("exchange_rate"))),
                })
            remaining = 0.0
        else:
            # 잔량이 작으면 다음 배치에서만 소진
            take = min(r, float(row_next["quantity"]))
            if take > 0:
                allocations.append({
                    "in_date": row_next["in_date"],
                    "taken_qty": take,
                    "unit_price": (None if pd.isna(row_next.get("unit_price")) else float(row_next.get("unit_price"))),
                    "exchange_rate": (None if pd.isna(row_next.get("exchange_rate")) else float(row_next.get("exchange_rate"))),
                })
            remaining -= take

    # 6) 요약 출력
    total = sum(a["taken_qty"] for a in allocations)
    shortage = max(0.0, float(sales_qty) - total)
    # 날짜를 문자열로 깔끔하게
    for a in allocations:
        if isinstance(a["in_date"], (pd.Timestamp, )):
            a["in_date"] = a["in_date"].strftime("%Y-%m-%d")
    print(f"[결과] {product} | 모드=LIFO(equal-lot) | 할당건수={len(allocations)} | 합계={total} | 기대값={sales_qty} | 부족분={shortage}")
    return allocations


# ======================
# 최종 실행
# ======================

def filter_and_fifo(result: dict):
    file_path = Path(result["file_path"])
    year = int(result["year"])

    # 기간 정보 추출
    start = pd.Timestamp(f"{year}-{int(result['start_month']):02d}-{int(result['start_day']):02d}")
    end = pd.Timestamp(f"{year}-{int(result['end_month']):02d}-{int(result['end_day']):02d}")

    sales_df = filter_excel_by_ui(result)
    final_results = []
    for _, row in sales_df.iterrows():
        product = row["product"]
        qty = row["quantity"]
        print(f"DEBUG - FIFO 처리중: {product}, 판매수량={qty}")
        # start와 cutoff 매개변수 추가
        allocations = fifo_for_product(file_path, product, year, qty, start=start, cutoff=end)
        print("DEBUG - FIFO 결과:", allocations)
        final_results.append({
            "product": product,
            "sales_qty": qty,
            "fifo_allocations": allocations
        })

    return final_results


# DEBUG - FIFO 처리중: AMBN, 판매수량=31200.0
# [처리중] AMBN | 모드=LIFO(equal-lot) | B=16000 | 판매수량=31200.0 | cutoff=2025-08-25
# [결과] AMBN | 모드=LIFO(equal-lot) | 할당건수=3 | 합계=31200.0 | 기대값=31200.0 | 부족분=0.0
# DEBUG - FIFO 결과: [{'in_date': '2025-08-14', 'taken_qty': 16000.0, 'unit_price': 10.85, 'exchange_rate': 1388.1}, {'in_date': '2025-07-29', 'taken_qty': 14400.0, 'unit_price': 10.85, 'exchange_rate': 1372.0}, {'in_date': '2025-07-21', 'taken_qty': 800.0, 'unit_price': 10.85, 'exchange_rate': 1391.8}]