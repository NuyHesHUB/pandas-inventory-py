import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd

from sw_exel_py_project.exel_reader import (
    _extract_year_from_sheet_name,
    _is_krw_number_format,
    _load_inbound_events_progressive,
    _pick_optional_col,
    _years_to_scan_desc,
    extract_inbound_events,
    fifo_for_product,
)


class ExelReaderTest(unittest.TestCase):
    def test_extract_year_from_sheet_name(self):
        self.assertEqual(_extract_year_from_sheet_name("재고관리대장 2025"), 2025)
        self.assertEqual(_extract_year_from_sheet_name("2024_기초재고"), 2024)
        self.assertIsNone(_extract_year_from_sheet_name("요약"))

    def test_years_to_scan_desc(self):
        with patch("sw_exel_py_project.exel_reader._available_inventory_years", return_value=(2023, 2024, 2025, 2026)):
            years = _years_to_scan_desc("dummy.xlsx", 2025)
        self.assertEqual(years, (2025, 2024, 2023))

    def test_inbound_progressive_backtracks_only_when_needed(self):
        base_df = pd.DataFrame(
            {
                "__product__": ["AMBN"],
                "__in_date__": [pd.Timestamp("2025-07-29")],
                "__in_qty__": [3000.0],
                "__unit_price__": [10.85],
                "__exchange_rate__": [1372.0],
                "__supplier__": [""],
                "__customer__": [""],
                "__note__": [""],
            }
        )
        prev_df = pd.DataFrame(
            {
                "__product__": ["AMBN"],
                "__in_date__": [pd.Timestamp("2024-12-20")],
                "__in_qty__": [5000.0],
                "__unit_price__": [10.5],
                "__exchange_rate__": [1365.0],
                "__supplier__": [""],
                "__customer__": [""],
                "__note__": [""],
            }
        )

        with patch("sw_exel_py_project.exel_reader._years_to_scan_desc", return_value=(2025, 2024, 2023)), \
             patch("sw_exel_py_project.exel_reader.load_inventory_df", side_effect=[prev_df]):
            out = _load_inbound_events_progressive(
                file_path=Path("dummy.xlsx"),
                year=2025,
                product="AMBN",
                cutoff=pd.Timestamp("2025-08-31"),
                base_df=base_df,
                min_required_qty=7000.0,
            )
        self.assertAlmostEqual(float(out["quantity"].sum()), 8000.0)
        self.assertIn(pd.Timestamp("2024-12-20"), set(pd.to_datetime(out["event_date"])))

    def test_pick_optional_col_matches_multiindex_subheader(self):
        cols = pd.MultiIndex.from_tuples(
            [
                ("입  고  (수  입)", "공 급 처"),
                ("납  품  (매  출)", "매 출 처"),
            ]
        )
        df = pd.DataFrame([["A", "B"]], columns=cols)
        self.assertEqual(_pick_optional_col(df, "공 급 처"), ("입  고  (수  입)", "공 급 처"))
        self.assertEqual(_pick_optional_col(df, "매 출 처"), ("납  품  (매  출)", "매 출 처"))

    def test_extract_inbound_events_excludes_return_rows(self):
        df = pd.DataFrame(
            {
                "__product__": ["CZ(P)", "CZ(P)"],
                "__in_date__": [pd.Timestamp("2025-01-07"), pd.Timestamp("2025-07-29")],
                "__in_qty__": [480.0, 3000.0],
                "__unit_price__": [2.28, 2.28],
                "__exchange_rate__": [None, None],
                "__supplier__": ["대진화공(반품)", "Puyang Willing"],
                "__customer__": ["", ""],
                "__note__": ["", ""],
            }
        )

        out = extract_inbound_events(
            file_path=Path("dummy.xlsx"),
            year=2025,
            product="CZ(P)",
            cutoff=pd.Timestamp("2025-08-31"),
            df=df,
        )

        self.assertEqual(len(out), 1)
        self.assertEqual(str(out.iloc[0]["event_date"].date()), "2025-07-29")
        self.assertEqual(float(out.iloc[0]["quantity"]), 3000.0)

    def test_extract_inbound_events_excludes_high_price_and_korean_supplier(self):
        df = pd.DataFrame(
            {
                "__product__": ["AMBN", "AMBN", "AMBN"],
                "__in_date__": [
                    pd.Timestamp("2025-07-10"),
                    pd.Timestamp("2025-07-11"),
                    pd.Timestamp("2025-07-12"),
                ],
                "__in_qty__": [1000.0, 2000.0, 3000.0],
                "__unit_price__": [10.85, 500.0, 11.1],
                "__exchange_rate__": [1370.0, 1380.0, 1390.0],
                "__supplier__": ["Puyang Willing", "Global Corp", "한국상사"],
                "__customer__": ["", "", ""],
                "__note__": ["", "", ""],
            }
        )

        out = extract_inbound_events(
            file_path=Path("dummy.xlsx"),
            year=2025,
            product="AMBN",
            cutoff=pd.Timestamp("2025-08-31"),
            df=df,
        )

        self.assertEqual(len(out), 1)
        self.assertEqual(str(out.iloc[0]["event_date"].date()), "2025-07-10")
        self.assertEqual(float(out.iloc[0]["unit_price"]), 10.85)

    def test_krw_number_format_detection(self):
        self.assertTrue(_is_krw_number_format('_-[$₩-412]* #,##0_-'))
        self.assertTrue(_is_krw_number_format('"원"#,##0'))
        self.assertFalse(_is_krw_number_format('$#,##0.00'))
        self.assertFalse(_is_krw_number_format('General'))

    def test_extract_inbound_events_excludes_krw_accounting_format(self):
        df = pd.DataFrame(
            {
                "__product__": ["AMBN", "AMBN"],
                "__in_date__": [pd.Timestamp("2025-07-10"), pd.Timestamp("2025-07-11")],
                "__in_qty__": [1000.0, 2000.0],
                "__unit_price__": [10.85, 11.1],
                "__exchange_rate__": [1370.0, 1380.0],
                "__supplier__": ["Puyang Willing", "Global Corp"],
                "__customer__": ["", ""],
                "__note__": ["", ""],
                "__is_krw_accounting_format__": [False, True],
                "__is_high_krw_unit_price__": [False, True],
            }
        )

        out = extract_inbound_events(
            file_path=Path("dummy.xlsx"),
            year=2025,
            product="AMBN",
            cutoff=pd.Timestamp("2025-08-31"),
            df=df,
        )

        self.assertEqual(len(out), 1)
        self.assertEqual(str(out.iloc[0]["event_date"].date()), "2025-07-10")

    def test_fifo_adjustment_seed_uses_real_inbound_date_not_boundary(self):
        product = "CZ(P)"
        start = pd.Timestamp("2025-07-23")
        cutoff = pd.Timestamp("2025-07-31")

        df = pd.DataFrame(
            {
                "__product__": [product, product, product],
                "__event_date__": [
                    pd.Timestamp("2025-06-30"),
                    pd.Timestamp("2025-07-23"),
                    pd.Timestamp("2025-07-29"),
                ],
                "__in_date__": [
                    pd.Timestamp("2025-06-30"),
                    pd.Timestamp("2025-07-23"),
                    pd.Timestamp("2025-07-29"),
                ],
                "__row_date__": [
                    pd.Timestamp("2025-06-30"),
                    pd.Timestamp("2025-07-23"),
                    pd.Timestamp("2025-07-29"),
                ],
                "__stock__": [100.0, 0.0, 3000.0],
                "__sales_qty__": [0.0, 100.0, 0.0],
                "__in_qty__": [0.0, 0.0, 3000.0],
                "__unit_price__": [None, None, 2.28],
                "__exchange_rate__": [None, None, None],
                "__supplier__": ["", "", "Puyang Willing"],
                "__customer__": ["", "", ""],
                "__note__": ["", "", ""],
            }
        )

        allocs = fifo_for_product(
            file_path=Path("dummy.xlsx"),
            product=product,
            year=2025,
            sales_qty=100.0,
            start=start,
            cutoff=cutoff,
            df=df,
        )

        self.assertTrue(allocs)
        self.assertEqual(str(pd.Timestamp(allocs[0]["in_date"]).date()), "2025-07-29")
        self.assertAlmostEqual(float(allocs[0]["taken_qty"]), 100.0)

    def test_fifo_negative_delta_assigns_to_oldest_prestart_lot(self):
        product = "AMBN"
        start = pd.Timestamp("2026-01-28")
        cutoff = pd.Timestamp("2026-02-19")

        # 2026 시트에 2025-12 이력이 일부만 존재하는 상황을 모사.
        # 시작 전 판매 집계(1/14)만으로는 장부재고(1/27, 26,720)보다 lot 잔량이 3,040 부족해진다.
        # 이 부족분은 FIFO 일관성을 위해 pre-start 최초 lot(2025-12-15)에 붙어야 한다.
        df = pd.DataFrame(
            {
                "__product__": [product] * 8,
                "__event_date__": [
                    pd.Timestamp("2025-12-15"),  # inbound
                    pd.Timestamp("2026-01-14"),  # pre-sale
                    pd.Timestamp("2026-01-19"),  # inbound
                    pd.Timestamp("2026-01-28"),  # sale
                    pd.Timestamp("2026-01-30"),  # sale
                    pd.Timestamp("2026-02-04"),  # sale
                    pd.Timestamp("2026-02-19"),  # inbound
                    pd.Timestamp("2026-02-19"),  # sale
                ],
                "__in_date__": [
                    pd.Timestamp("2025-12-15"),
                    pd.Timestamp("2026-01-14"),
                    pd.Timestamp("2026-01-19"),
                    pd.Timestamp("2026-01-28"),
                    pd.Timestamp("2026-01-30"),
                    pd.Timestamp("2026-02-04"),
                    pd.Timestamp("2026-02-19"),
                    pd.Timestamp("2026-02-19"),
                ],
                "__row_date__": [
                    pd.Timestamp("2025-12-15"),
                    pd.Timestamp("2026-01-14"),
                    pd.Timestamp("2026-01-19"),
                    pd.Timestamp("2026-01-28"),
                    pd.Timestamp("2026-01-30"),
                    pd.Timestamp("2026-02-04"),
                    pd.Timestamp("2026-02-19"),
                    pd.Timestamp("2026-02-19"),
                ],
                "__stock__": [
                    35680.0,
                    10720.0,
                    26720.0,  # boundary(1/27) 기준 마지막 재고
                    18400.0,
                    10080.0,
                    1760.0,
                    25760.0,
                    17440.0,
                ],
                "__sales_qty__": [0.0, 8320.0, 0.0, 8320.0, 8320.0, 8320.0, 0.0, 8320.0],
                "__in_qty__": [16000.0, 0.0, 16000.0, 0.0, 0.0, 0.0, 24000.0, 0.0],
                "__unit_price__": [10.74, None, 10.74, None, None, None, 10.74, None],
                "__exchange_rate__": [1475.0, None, 1471.5, None, None, None, 1457.4, None],
                "__supplier__": ["", "", "", "", "", "", "", ""],
                "__customer__": ["", "", "", "", "", "", "", ""],
                "__note__": ["", "", "", "", "", "", "", ""],
            }
        )

        allocs = fifo_for_product(
            file_path=Path("dummy.xlsx"),
            product=product,
            year=2026,
            sales_qty=33280.0,
            start=start,
            cutoff=cutoff,
            df=df,
        )

        by_date = {str(a["in_date"]): float(a["taken_qty"]) for a in allocs}
        self.assertAlmostEqual(sum(by_date.values()), 33280.0)
        self.assertAlmostEqual(by_date["2025-12-15"], 10720.0)
        self.assertAlmostEqual(by_date["2026-01-19"], 16000.0)
        self.assertAlmostEqual(by_date["2026-02-19"], 6560.0)

    def test_fifo_shortage_adjustment_preserves_sales_quantity(self):
        product = "Cobalt Tetroxide"
        start = pd.Timestamp("2025-08-01")
        cutoff = pd.Timestamp("2025-08-31")

        df = pd.DataFrame(
            {
                "__product__": [product, product],
                "__event_date__": [pd.Timestamp("2025-07-01"), pd.Timestamp("2025-08-21")],
                "__in_date__": [pd.Timestamp("2025-07-01"), pd.Timestamp("2025-08-21")],
                "__row_date__": [pd.Timestamp("2025-07-01"), pd.Timestamp("2025-08-21")],
                "__stock__": [2000.0, 0.0],
                "__sales_qty__": [0.0, 2300.0],
                "__in_qty__": [2000.0, 0.0],
                "__unit_price__": [12.5, None],
                "__exchange_rate__": [1370.0, None],
                "__supplier__": ["Global Supplier", ""],
                "__customer__": ["", ""],
                "__note__": ["", ""],
            }
        )

        with patch("sw_exel_py_project.exel_reader._years_to_scan_desc", return_value=(2025,)):
            allocs = fifo_for_product(
                file_path=Path("dummy.xlsx"),
                product=product,
                year=2025,
                sales_qty=2300.0,
                start=start,
                cutoff=cutoff,
                df=df,
            )

        self.assertAlmostEqual(sum(float(a["taken_qty"]) for a in allocs), 2300.0)
        self.assertEqual(str(pd.Timestamp(allocs[0]["in_date"]).date()), "2025-07-01")
        self.assertAlmostEqual(float(allocs[0]["unit_price"]), 12.5)


if __name__ == "__main__":
    unittest.main()
