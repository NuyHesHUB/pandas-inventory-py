import unittest
from pathlib import Path
from unittest.mock import patch

import openpyxl
import pandas as pd

from sw_exel_py_project.cost_report import (
    TemplateRow,
    _build_alias_map,
    _match_inbound_qty,
    _parse_note_date,
    generate_cost_report_without_template,
)


class CostReportTest(unittest.TestCase):
    def test_parse_note_date_excel_serial(self):
        dt = _parse_note_date(45825)
        self.assertIsNotNone(dt)
        self.assertEqual(str(dt.date()), "2025-06-17")

    def test_parse_note_date_text(self):
        dt = _parse_note_date("2025/07/29 입고(경동)")
        self.assertIsNotNone(dt)
        self.assertEqual(str(dt.date()), "2025-07-29")

    def test_alias_map_canonical_match(self):
        template_items = ["CG-172-25KG", "MOCA(China)"]
        source_products = ["CG 172-25KG", "MOCA"]
        alias = _build_alias_map(template_items, source_products)
        self.assertEqual(alias["CG-172-25KG"], "CG 172-25KG")
        self.assertEqual(alias["MOCA(China)"], "MOCA")

    def test_match_inbound_qty(self):
        inbound_df = pd.DataFrame(
            {
                "product": ["AMBN", "AMBN"],
                "in_date": [pd.Timestamp("2025-07-29"), pd.Timestamp("2025-08-14")],
                "quantity": [16000.0, 16000.0],
                "unit_price": [10.85, 10.85],
                "exchange_rate": [1372.0, 1388.1],
            }
        )
        row = TemplateRow(
            excel_row=12,
            item="AMBN",
            unit_price_raw=10.85,
            unit_price_num=10.85,
            exchange_rate=1372.0,
            duty_rate=0.0,
            duty_factor=0.0,
            quantity=0.0,
            note="2025/07/29",
            note_date=pd.Timestamp("2025-07-29"),
        )
        qty = _match_inbound_qty(inbound_df, "AMBN", row)
        self.assertEqual(qty, 16000.0)

    def test_generate_cost_report_without_template(self):
        df = pd.DataFrame(
            {
                "__event_date__": [pd.Timestamp("2026-01-10")],
                "__product__": ["AMBN"],
                "__sales_qty__": [100.0],
            }
        )
        allocs = [
            {
                "in_date": "2026-01-19",
                "taken_qty": 100.0,
                "unit_price": 10.74,
                "exchange_rate": 1471.5,
            }
        ]

        td_path = Path("artifacts") / "tmp_direct_report_test"
        td_path.mkdir(parents=True, exist_ok=True)
        inventory_file = td_path / "inventory.xlsx"
        output_file = td_path / "out.xlsx"
        inventory_file.write_bytes(b"")

        with patch("sw_exel_py_project.cost_report.load_inventory_df", return_value=df), \
             patch("sw_exel_py_project.cost_report.fifo_for_product", return_value=allocs):
            summary = generate_cost_report_without_template(
                inventory_file=inventory_file,
                output_file=output_file,
                year=2026,
                start=pd.Timestamp("2026-01-01"),
                end=pd.Timestamp("2026-01-31"),
            )

        self.assertTrue(output_file.exists())
        self.assertEqual(summary["template_rows"], 1)
        self.assertEqual(summary["template_items"], 1)
        self.assertEqual(summary["fallback_items"], [])

        wb = openpyxl.load_workbook(output_file)
        ws = wb[wb.sheetnames[0]]
        self.assertEqual(ws.cell(2, 1).value, "AMBN")
        self.assertEqual(float(ws.cell(2, 2).value), 10.74)
        self.assertEqual(float(ws.cell(2, 3).value), 1471.5)
        self.assertEqual(float(ws.cell(2, 6).value), 100.0)


if __name__ == "__main__":
    unittest.main()
