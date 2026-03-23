from pathlib import Path
import pandas as pd
from sw_exel_py_project.exel_reader import load_inventory_df, fifo_for_product, sum_sales_in_range

# Adjust these if your actual Excel filename or window differs
ROOT = Path(r"c:\Users\sehye\바탕 화면\code\project\sw-exel-py-project")
XLSX = ROOT / "test.xlsx"
YEAR = 2025
PRODUCT = "AMBN"
START = pd.Timestamp(f"{YEAR}-07-24")
END = pd.Timestamp(f"{YEAR}-08-25")

df = load_inventory_df(str(XLSX), YEAR)
S = sum_sales_in_range(XLSX, YEAR, PRODUCT, START, END, df=df)
allocs = fifo_for_product(XLSX, PRODUCT, YEAR, S, start=START, cutoff=END, df=df)

print(f"판매합계: {S}")
for a in allocs:
    print(f"입고일: {a['in_date']}, 소진량: {a['taken_qty']}, 단가: {a.get('unit_price')}, 환율: {a.get('exchange_rate')}")
