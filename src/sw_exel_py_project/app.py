"""UI 결과를 받아 원가 결과물을 생성하는 공통 실행 로직.

CLI(__main__)와 데스크톱(desktop) 진입점이 같은 코드를 쓰도록 분리했다.
"""

from pathlib import Path

import pandas as pd

from .cost_report import generate_cost_report_without_template, renew_cost_report_from_template


def run_cost_report_from_ui_result(result: dict) -> tuple[dict, bool]:
    """run_ui()가 돌려준 dict로 원가 결과물을 생성하고 (summary, use_template)을 반환한다."""
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
    else:
        summary = generate_cost_report_without_template(
            inventory_file=Path(result["file_path"]),
            output_file=Path(result["output_file"]),
            year=year,
            start=start,
            end=end,
        )
    return summary, use_template


def format_summary_lines(summary: dict) -> list[str]:
    """실행 결과 요약을 사람이 읽을 문장 목록으로 만든다. (콘솔/메시지박스 공용)"""
    lines = [
        f"저장 위치: {summary['output_file']}",
        f"데이터 행 수: {summary['template_rows']}",
        f"품목 수: {summary['template_items']}",
        f"매핑 성공 품목 수: {summary.get('mapped_items', 0)}",
        f"변경 행 수: {summary.get('changed_rows', 0)}",
        f"변경 셀 수: {summary.get('changed_cells', 0)}",
    ]
    if summary.get("output_redirected"):
        lines.append("안내: 템플릿과 같은 저장 경로가 입력되어 자동으로 새 파일명으로 저장했습니다.")
    if summary.get("unresolved_items"):
        lines.append("매핑 실패 품목: " + ", ".join(summary["unresolved_items"]))
    if summary.get("fallback_items"):
        lines.append("단가 lot 미탐지(기존값 사용) 품목: " + ", ".join(summary["fallback_items"]))
    if summary.get("excluded_price_items"):
        lines.append("원화/국내 단가 제외 품목: " + ", ".join(summary["excluded_price_items"]))
    if summary.get("anchored_items"):
        lines.append("기존 템플릿 lot 유지 품목: " + ", ".join(summary["anchored_items"]))
    if summary.get("overflow_items"):
        lines.append("lot 행수 초과(템플릿 행 자동 추가) 품목: " + ", ".join(summary["overflow_items"]))
    if summary.get("changed_cells", 0) == 0 and summary.get("anchored_items"):
        lines.append("안내: 선택 기간과 템플릿 기간이 같아 기존 템플릿 lot을 보존했습니다.")
    elif summary.get("changed_cells", 0) == 0:
        lines.append("경고: 갱신된 셀이 없습니다. 매핑 실패 또는 기간/데이터 조건을 확인하세요.")
    return lines
