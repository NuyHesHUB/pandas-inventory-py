import argparse
from pathlib import Path
import pandas as pd

from .app import format_summary_lines, run_cost_report_from_ui_result
from .cost_report import generate_cost_report_without_template, renew_cost_report_from_template
from .ui.main_ui import run_ui


def _parse_date_arg(s: str) -> pd.Timestamp:
    return pd.to_datetime(s, format="%Y-%m-%d", errors="raise")

def main():
    p = argparse.ArgumentParser(description="Read a single column from an Excel file and print values.")
    p.add_argument("--file", help="Path to .xlsx file")
    p.add_argument("--col", help="Column name (header) OR Excel letter like A,B,AA")
    p.add_argument("--col-index", type=int, help="Zero-based column index (0,1,2,...)")
    p.add_argument("--sheet", default=0, help="Sheet name or index (default 0)")
    p.add_argument("--no-header", action="store_true", help="Treat first row as data (no header)")
    p.add_argument("--head", type=int, default=0, help="Print only first N rows (0 = all)")
    p.add_argument("--ui", action="store_true", help="Show UI for file and date selection")

    # 결과물 템플릿 자동 갱신
    p.add_argument("--renew-cost-report", action="store_true", help="Renew cost report by mapping inventory sales into result template.")
    p.add_argument("--inventory-file", help="Path to inventory workbook (.xlsx)")
    p.add_argument("--template-file", help="Path to result template workbook (.xlsx)")
    p.add_argument("--output-file", help="Path to save renewed workbook (.xlsx)")
    p.add_argument("--year", type=int, help="Target year (e.g. 2025)")
    p.add_argument("--start-date", help="Start date in YYYY-MM-DD")
    p.add_argument("--end-date", help="End date in YYYY-MM-DD")
    p.add_argument("--sheet-name", help="Template sheet name (default: first sheet)")
    p.add_argument("--template-mode", choices=["on", "off"], default="on", help="Template usage mode: on=use template, off=generate workbook without template.")

    args = p.parse_args()

    if args.renew_cost_report:
        use_template = args.template_mode == "on"
        required = {
            "--inventory-file": args.inventory_file,
            "--output-file": args.output_file,
            "--year": args.year,
            "--start-date": args.start_date,
            "--end-date": args.end_date,
        }
        if use_template:
            required["--template-file"] = args.template_file
        missing = [k for k, v in required.items() if v in (None, "")]
        if missing:
            print("필수 옵션 누락:", ", ".join(missing))
            return

        start = _parse_date_arg(args.start_date)
        end = _parse_date_arg(args.end_date)

        if use_template:
            summary = renew_cost_report_from_template(
                inventory_file=Path(args.inventory_file),
                template_file=Path(args.template_file),
                output_file=Path(args.output_file),
                year=int(args.year),
                start=start,
                end=end,
                sheet_name=args.sheet_name,
            )
            print("=== 결과물 템플릿 갱신 완료 ===")
        else:
            summary = generate_cost_report_without_template(
                inventory_file=Path(args.inventory_file),
                output_file=Path(args.output_file),
                year=int(args.year),
                start=start,
                end=end,
                sheet_name=args.sheet_name,
            )
            print("=== 결과물 신규 생성 완료(템플릿 미사용) ===")

        for line in format_summary_lines(summary):
            print(line)
        return

    if args.ui:
        def _runner(result: dict) -> list[str]:
            summary, use_template = run_cost_report_from_ui_result(result)
            if use_template:
                print("=== 결과물 템플릿 갱신 완료(UI) ===")
            else:
                print("=== 결과물 신규 생성 완료(UI, 템플릿 미사용) ===")
            lines = format_summary_lines(summary)
            for line in lines:
                print(line)
            return lines

        run_ui(on_run=_runner)
        return

    if not args.file:
        print("엑셀 파일 경로를 --file 옵션으로 지정하세요.")
        return



if __name__ == "__main__":
    main()
