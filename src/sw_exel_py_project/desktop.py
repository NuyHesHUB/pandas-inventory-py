"""데스크톱(더블클릭 실행)용 진입점.

PyInstaller --windowed 빌드에는 콘솔이 없으므로, 처리 결과와 오류를
stdout 대신 tkinter 메시지박스로 보여준다.
"""

from .app import format_summary_lines, run_cost_report_from_ui_result
from .ui.main_ui import run_ui


def _show_message(kind: str, title: str, text: str) -> None:
    import tkinter as tk
    from tkinter import messagebox

    root = tk.Tk()
    root.withdraw()
    if kind == "error":
        messagebox.showerror(title, text)
    else:
        messagebox.showinfo(title, text)
    root.destroy()


def run() -> None:
    result = run_ui()
    if not result:
        return

    try:
        summary, use_template = run_cost_report_from_ui_result(result)
    except Exception as e:  # 콘솔이 없으니 어떤 실패든 사용자에게 보여야 한다.
        _show_message("error", "원가 결과물 생성 실패", str(e))
        return

    header = "결과물 템플릿 갱신 완료" if use_template else "결과물 신규 생성 완료 (템플릿 미사용)"
    _show_message("info", "원가 결과물 생성", header + "\n\n" + "\n".join(format_summary_lines(summary)))


if __name__ == "__main__":
    run()
