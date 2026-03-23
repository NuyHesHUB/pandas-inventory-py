import argparse
from pathlib import Path
import pandas as pd

from .cost_report import generate_cost_report_without_template, renew_cost_report_from_template
from .exel_reader import filter_and_fifo, filter_excel_by_ui
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

        print("저장 위치:", summary["output_file"])
        print("데이터 행 수:", summary["template_rows"])
        print("품목 수:", summary["template_items"])
        print("매핑 성공 품목 수:", summary.get("mapped_items", 0))
        print("변경 행 수:", summary.get("changed_rows", 0))
        print("변경 셀 수:", summary.get("changed_cells", 0))
        if summary.get("output_redirected"):
            print("안내: 템플릿과 같은 저장 경로가 입력되어 자동으로 새 파일명으로 저장했습니다.")
        if summary["unresolved_items"]:
            print("매핑 실패 품목:", ", ".join(summary["unresolved_items"]))
        if summary.get("fallback_items"):
            print("단가 lot 미탐지(기존값 사용) 품목:", ", ".join(summary["fallback_items"]))
        if summary.get("overflow_items"):
            print("lot 행수 초과(마지막 행 합산) 품목:", ", ".join(summary["overflow_items"]))
        if summary.get("changed_cells", 0) == 0:
            print("경고: 갱신된 셀이 없습니다. 매핑 실패 또는 기간/데이터 조건을 확인하세요.")
        return

    if args.ui:
        result = run_ui()
        if not result:
            print("실행이 취소되었습니다.")
            return

        if result.get("mode") == "cost_report":
            year = int(result["year"])
            start = pd.Timestamp(f"{year}-{int(result['start_month']):02d}-{int(result['start_day']):02d}")
            end = pd.Timestamp(f"{year}-{int(result['end_month']):02d}-{int(result['end_day']):02d}")
            use_template = bool(result.get("use_template", True))

            if use_template:
                summary = renew_cost_report_from_template(
                    inventory_file=Path(result["file_path"]),
                    template_file=Path(result["template_file"]),
                    output_file=Path(result["output_file"]),
                    year=year,
                    start=start,
                    end=end,
                )
                print("=== 결과물 템플릿 갱신 완료(UI) ===")
            else:
                summary = generate_cost_report_without_template(
                    inventory_file=Path(result["file_path"]),
                    output_file=Path(result["output_file"]),
                    year=year,
                    start=start,
                    end=end,
                )
                print("=== 결과물 신규 생성 완료(UI, 템플릿 미사용) ===")

            print("저장 위치:", summary["output_file"])
            print("데이터 행 수:", summary["template_rows"])
            print("품목 수:", summary["template_items"])
            print("매핑 성공 품목 수:", summary.get("mapped_items", 0))
            print("변경 행 수:", summary.get("changed_rows", 0))
            print("변경 셀 수:", summary.get("changed_cells", 0))
            if summary.get("output_redirected"):
                print("안내: 템플릿과 같은 저장 경로가 입력되어 자동으로 새 파일명으로 저장했습니다.")
            if summary["unresolved_items"]:
                print("매핑 실패 품목:", ", ".join(summary["unresolved_items"]))
            if summary.get("fallback_items"):
                print("단가 lot 미탐지(기존값 사용) 품목:", ", ".join(summary["fallback_items"]))
            if summary.get("overflow_items"):
                print("lot 행수 초과(마지막 행 합산) 품목:", ", ".join(summary["overflow_items"]))
            if summary.get("changed_cells", 0) == 0:
                print("경고: 갱신된 셀이 없습니다. 매핑 실패 또는 기간/데이터 조건을 확인하세요.")
            return

        pd.set_option('display.max_rows', None)

        # 1) 판매 합계 DataFrame (원래 출력하던 것)
        sales_df = filter_excel_by_ui(result)
        print("=== 판매 합계 ===")
        print(sales_df)

        # 2) FIFO 추적 결과
        fifo_results = filter_and_fifo(result)
        print("\n=== FIFO 추적 결과 ===")
        for item in fifo_results:
            print(f"품목: {item['product']}, 판매합계: {item['sales_qty']}")
            for alloc in item["fifo_allocations"]:
                print(
                    f"  입고일: {alloc['in_date']}, "
                    f"소진량: {alloc['taken_qty']}, "
                    f"단가: {alloc.get('unit_price')}, "
                    f"환율: {alloc.get('exchange_rate')}"
                )
        return

    if not args.file:
        print("엑셀 파일 경로를 --file 옵션으로 지정하세요.")
        return



if __name__ == "__main__":
    main()
