"""데스크톱(더블클릭 실행)용 진입점.

UI가 실행/진행/완료 표시와 오류 메시지박스까지 담당하므로,
여기서는 결과물 생성 로직만 연결한다.
"""

from .ui.main_ui import run_ui


def _runner(result: dict) -> list[str]:
    # pandas 임포트 비용(수백 ms)을 실행 시점으로 미뤄 창이 먼저 뜨게 한다.
    from .app import format_summary_lines, run_cost_report_from_ui_result

    summary, _use_template = run_cost_report_from_ui_result(result)
    return format_summary_lines(summary)


def run() -> None:
    run_ui(on_run=_runner)


if __name__ == "__main__":
    run()
