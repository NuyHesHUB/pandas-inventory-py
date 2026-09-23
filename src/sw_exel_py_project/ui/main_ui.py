"""원가 결과물 생성 UI.

claude.ai/design '원가 결과물 생성 v2' 디자인을 tkinter로 옮긴 화면.
타이틀바는 OS 네이티브를 쓰고, 본문 레이아웃/색/상태바만 디자인을 따른다.

입력 정책:
- 재고관리대장 / 저장 경로 모두 필수 (템플릿 없이 항상 신규 생성 모드)
- 연도는 실행 시점의 올해가 기본값, 기간(시작~종료)은 달력에서 직접 선택
"""

import calendar as _calendar
import sys
import threading
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk
from datetime import date, timedelta
from math import ceil
from pathlib import Path

from .app_icon import APP_ICON_PNG_BASE64

# 디자인 팔레트 (accent = oklch(0.5 0.14 250) 근사값)
ACCENT = "#0465af"
ACCENT_DARK = "#03528f"
ACCENT_LIGHT = "#c7dbf2"
ACCENT_FAINT = "#edf3fa"
BG = "#f3f3f3"
CARD_BG = "#fbfbfb"
CARD_BORDER = "#e3e3e3"
FIELD_BG = "#ffffff"
FIELD_BORDER = "#c9c9c9"
BTN_BG = "#fdfdfd"
BTN_HOVER = "#f5f5f5"
BTN_BORDER = "#b8b8b8"
FOOTER_BG = "#ebebeb"
FOOTER_BORDER = "#e0e0e0"
TEXT = "#1f1f1f"
MUTED = "#5a5a5a"
PLACEHOLDER = "#8a8a8a"
RED = "#c0392b"
GREEN = "#2e7d4f"
DISABLED_BG = "#e0e0e0"
DISABLED_FG = "#9a9a9a"

_WEEKDAY_KR = ["월", "화", "수", "목", "금", "토", "일"]


def run_ui(on_run=None):
    """UI를 띄운다.

    on_run이 주어지면 실행 버튼이 창을 닫지 않고 백그라운드 스레드로
    on_run(result)을 호출한 뒤 결과를 상태바/메시지박스로 보여준다.
    on_run이 None이면 기존처럼 result dict를 반환하고 창을 닫는다.
    """
    state = {
        "source": "",
        "output": "",
        "start": None,  # (month, day)
        "end": None,
        "open": None,  # "start" | "end"
        "view_month": date.today().month,
        "running": False,
        "done_file": "",
    }
    legacy_result = {}

    root = tk.Tk()
    root.title("원가 결과물 생성")
    root.configure(bg=BG)
    root.resizable(False, False)
    try:
        icon = tk.PhotoImage(data=APP_ICON_PNG_BASE64)
        root.iconphoto(True, icon)
        root._app_icon = icon  # PhotoImage가 GC로 사라지지 않게 참조 유지
    except tk.TclError:
        pass

    default_font = tkfont.nametofont("TkDefaultFont")
    if sys.platform == "win32":
        family, size = "Segoe UI", 10
    elif sys.platform == "darwin":
        family, size = default_font.actual("family"), default_font.actual("size")
    else:
        family, size = default_font.actual("family"), 10
    base = (family, size)
    bold = (family, size, "bold")
    small = (family, max(size - 1, 8))
    small_bold = (family, max(size - 1, 8), "bold")

    # ----- 공통 위젯 헬퍼 -----

    def flat_button(parent, text, command, padx=12, pady=4):
        btn = tk.Label(
            parent, text=text, font=base, bg=BTN_BG, fg=TEXT,
            padx=padx, pady=pady, highlightthickness=1, highlightbackground=BTN_BORDER,
        )
        btn.bind("<Enter>", lambda _e: btn.state_enabled and btn.config(bg=BTN_HOVER))
        btn.bind("<Leave>", lambda _e: btn.state_enabled and btn.config(bg=BTN_BG))
        btn.bind("<Button-1>", lambda _e: btn.state_enabled and command())
        btn.state_enabled = True
        return btn

    def field_label(parent):
        return tk.Label(
            parent, anchor="w", font=base, bg=FIELD_BG, fg=PLACEHOLDER,
            padx=8, highlightthickness=1, highlightbackground=FIELD_BORDER,
        )

    def card(parent, title):
        tk.Label(parent, text=title, font=small_bold, fg=MUTED, bg=BG).pack(anchor="w", padx=2)
        outer = tk.Frame(parent, bg=CARD_BG, highlightthickness=1, highlightbackground=CARD_BORDER)
        outer.pack(fill="x", pady=(4, 12))
        inner = tk.Frame(outer, bg=CARD_BG)
        inner.pack(fill="x", padx=14, pady=12)
        inner.columnconfigure(0, minsize=96)
        inner.columnconfigure(1, weight=1)
        return inner

    body = tk.Frame(root, bg=BG)
    body.pack(fill="both", expand=True, padx=16, pady=(12, 4))
    tk.Frame(body, width=590, height=1, bg=BG).pack()  # 창 폭 고정

    # ----- 파일 섹션 -----

    files = card(body, "파일")

    def pick_source():
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            state["source"] = path
            state["done_file"] = ""
            refresh()

    def pick_output():
        initial = "결과물_자동생성.xlsx"
        if state["source"]:
            initial = f"{Path(state['source']).stem}_원가결과_자동생성.xlsx"
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=initial,
        )
        if path:
            state["output"] = path
            state["done_file"] = ""
            refresh()

    rows = [
        ("재고관리대장", "source", pick_source),
        ("저장 위치", "output", pick_output),
    ]
    file_fields = {}
    for i, (label_text, key, cmd) in enumerate(rows):
        tk.Label(files, text=label_text, font=base, bg=CARD_BG, fg=TEXT).grid(
            row=i, column=0, sticky="w", pady=4
        )
        field = field_label(files)
        field.grid(row=i, column=1, sticky="ew", padx=(0, 10), pady=4, ipady=4)
        file_fields[key] = field
        flat_button(files, "찾아보기…", cmd).grid(row=i, column=2, pady=4)

    # ----- 기간 섹션 -----

    period = card(body, "기간")

    tk.Label(period, text="연도", font=base, bg=CARD_BG, fg=TEXT).grid(row=0, column=0, sticky="w", pady=4)
    year_var = tk.StringVar(value=str(date.today().year))
    year_box = tk.Spinbox(
        period, from_=2000, to=2100, textvariable=year_var, width=6, font=base,
        relief="flat", highlightthickness=1, highlightbackground=FIELD_BORDER,
        bg=FIELD_BG, buttonbackground=BTN_BG, command=lambda: on_year_changed(),
    )
    year_box.grid(row=0, column=1, sticky="w", pady=4, ipady=2)
    year_box.bind("<KeyRelease>", lambda _e: on_year_changed())

    tk.Label(period, text="시작 ~ 종료", font=base, bg=CARD_BG, fg=TEXT).grid(row=1, column=0, sticky="w", pady=4)
    dates_row = tk.Frame(period, bg=CARD_BG)
    dates_row.grid(row=1, column=1, columnspan=2, sticky="w", pady=4)

    def date_button(which):
        wrap = tk.Frame(dates_row, bg=FIELD_BG, highlightthickness=1, highlightbackground=FIELD_BORDER)
        text = tk.Label(wrap, font=base, bg=FIELD_BG, fg=PLACEHOLDER, anchor="w", padx=8, width=14)
        text.pack(side="left", ipady=4)
        arrow = tk.Label(wrap, text="▾", font=small, bg=FIELD_BG, fg=MUTED, padx=6)
        arrow.pack(side="left")
        for w in (wrap, text, arrow):
            w.bind("<Button-1>", lambda _e: toggle_picker(which))
        return wrap, text

    start_wrap, start_text = date_button("start")
    start_wrap.pack(side="left")
    tk.Label(dates_row, text=" ~ ", font=base, bg=CARD_BG, fg=MUTED).pack(side="left")
    end_wrap, end_text = date_button("end")
    end_wrap.pack(side="left")
    days_label = tk.Label(dates_row, text="", font=small, bg=CARD_BG, fg=MUTED)
    days_label.pack(side="left", padx=(10, 0))

    # ----- 달력 팝업 -----

    popup = tk.Frame(root, bg=FIELD_BG, highlightthickness=1, highlightbackground="#d0d0d0")

    def current_year():
        try:
            y = int(year_var.get())
        except ValueError:
            return None
        return y if 2000 <= y <= 2100 else None

    def to_date(pair):
        y = current_year()
        if y is None or pair is None:
            return None
        try:
            return date(y, pair[0], pair[1])
        except ValueError:
            return None

    def close_picker():
        state["open"] = None
        popup.place_forget()

    def toggle_picker(which):
        if state["running"]:
            return
        if state["open"] == which:
            close_picker()
            return
        state["open"] = which
        picked = state[which]
        state["view_month"] = picked[0] if picked else date.today().month
        render_calendar()

    def position_popup():
        """팝업이 창 밖으로 잘리지 않게 필드 아래/위 중 들어가는 쪽에 붙인다."""
        anchor = start_wrap if state["open"] == "start" else end_wrap
        popup.update_idletasks()
        height = popup.winfo_reqheight()
        anchor_bottom = anchor.winfo_rooty() - root.winfo_rooty() + anchor.winfo_height()
        if anchor_bottom + 4 + height <= root.winfo_height():
            popup.place(in_=anchor, relx=0, rely=1.0, y=4, anchor="nw")
        else:
            popup.place(in_=anchor, relx=0, rely=0, y=-4, anchor="sw")
        popup.lift()

    def move_month(delta):
        state["view_month"] = min(12, max(1, state["view_month"] + delta))
        render_calendar()

    def render_calendar():
        y = current_year()
        if y is None:
            close_picker()
            return
        for child in popup.winfo_children():
            child.destroy()

        head = tk.Frame(popup, bg=FIELD_BG)
        head.pack(fill="x", padx=10, pady=(8, 4))
        tk.Label(head, text=f"{y}년 {state['view_month']}월", font=bold, bg=FIELD_BG, fg=TEXT).pack(side="left")
        nav = tk.Frame(head, bg=FIELD_BG)
        nav.pack(side="right")
        for sym, delta in (("‹", -1), ("›", 1)):
            b = tk.Label(nav, text=sym, font=base, bg=FIELD_BG, fg=MUTED, width=3)
            b.pack(side="left")
            b.bind("<Button-1>", lambda _e, d=delta: move_month(d))
            b.bind("<Enter>", lambda _e, w=b: w.config(bg=BTN_HOVER))
            b.bind("<Leave>", lambda _e, w=b: w.config(bg=FIELD_BG))

        grid = tk.Frame(popup, bg=FIELD_BG)
        grid.pack(padx=10, pady=(0, 10))
        for col, name in enumerate(["일", "월", "화", "수", "목", "금", "토"]):
            tk.Label(
                grid, text=name, font=small, bg=FIELD_BG, width=4,
                fg=RED if col == 0 else PLACEHOLDER,
            ).grid(row=0, column=col)

        vm = state["view_month"]
        first_idx = (date(y, vm, 1).weekday() + 1) % 7  # 일요일 시작
        ndays = _calendar.monthrange(y, vm)[1]
        cells = ceil((first_idx + ndays) / 7) * 7
        grid_start = date(y, vm, 1) - timedelta(days=first_idx)

        sel = to_date(state[state["open"]]) if state["open"] else None
        s_d, e_d = to_date(state["start"]), to_date(state["end"])

        for i in range(cells):
            d = grid_start + timedelta(days=i)
            in_month = d.month == vm and d.year == y
            is_sel = sel is not None and d == sel
            is_end = (s_d is not None and d == s_d) or (e_d is not None and d == e_d)
            in_range = s_d is not None and e_d is not None and s_d < d < e_d

            bg_c = ACCENT if is_sel else ACCENT_LIGHT if is_end else ACCENT_FAINT if in_range else FIELD_BG
            fg_c = "#ffffff" if is_sel else "#bdbdbd" if not in_month else RED if d.weekday() == 6 else TEXT
            cell = tk.Label(
                grid, text=str(d.day), width=4, pady=3, bg=bg_c, fg=fg_c,
                font=small_bold if (is_sel or is_end) else small,
            )
            cell.grid(row=1 + i // 7, column=i % 7, pady=1)
            if d.year == y:
                cell.bind("<Button-1>", lambda _e, dd=d: pick_day(dd))

        position_popup()

    def pick_day(d):
        state[state["open"]] = (d.month, d.day)
        state["done_file"] = ""
        close_picker()
        refresh()

    def on_year_changed():
        # 연도가 바뀌어 무효가 된 날짜(예: 2월 29일)는 초기화
        for key in ("start", "end"):
            if state[key] is not None and to_date(state[key]) is None:
                state[key] = None
        state["done_file"] = ""
        if state["open"]:
            render_calendar()
        refresh()

    def on_global_click(event):
        if state["open"] is None:
            return
        w = str(event.widget)
        inside = w.startswith(str(popup)) or w.startswith(str(start_wrap)) or w.startswith(str(end_wrap))
        if not inside:
            close_picker()

    root.bind("<Button-1>", on_global_click, add="+")

    # ----- 하단 상태바 -----

    tk.Frame(root, height=1, bg=FOOTER_BORDER).pack(fill="x")
    footer = tk.Frame(root, bg=FOOTER_BG)
    footer.pack(fill="x")

    status_area = tk.Frame(footer, bg=FOOTER_BG)
    status_area.pack(side="left", fill="x", expand=True, padx=16, pady=12)
    status_label = tk.Label(status_area, font=small, bg=FOOTER_BG, fg=MUTED, anchor="w")
    status_label.pack(side="left")

    style = ttk.Style(root)
    style.theme_use("default")
    style.configure(
        "Accent.Horizontal.TProgressbar",
        background=ACCENT, troughcolor="#d4d4d4", borderwidth=0, thickness=4,
    )
    progress = ttk.Progressbar(
        status_area, mode="indeterminate", length=180, style="Accent.Horizontal.TProgressbar"
    )

    run_btn = tk.Label(
        footer, text="실행", font=base, padx=22, pady=4,
        bg=DISABLED_BG, fg=DISABLED_FG, highlightthickness=1, highlightbackground="#d4d4d4",
    )
    close_btn = flat_button(footer, "닫기", lambda: root.quit(), padx=18)
    run_btn.pack(side="right", padx=(0, 16), pady=12)
    close_btn.pack(side="right", padx=(0, 12), pady=12)

    # ----- 상태 계산/갱신 -----

    def fmt_pair(pair):
        d = to_date(pair)
        if d is None:
            return None
        return f"{d.month:02d}월 {d.day:02d}일 ({_WEEKDAY_KR[d.weekday()]})"

    def validation_status():
        """(메시지, 색상, 실행가능) 순서는 디자인의 상태바 로직을 따른다."""
        if not state["source"]:
            return "재고관리대장 파일을 선택하세요.", MUTED, False
        if not state["output"]:
            return "저장할 파일 경로를 선택하세요.", MUTED, False
        if Path(state["source"]).resolve(strict=False) == Path(state["output"]).resolve(strict=False):
            return "저장 경로가 재고관리대장 파일과 같습니다. 다른 경로를 선택하세요.", RED, False
        if current_year() is None:
            return "연도를 올바르게 입력하세요.", RED, False
        s_d, e_d = to_date(state["start"]), to_date(state["end"])
        if s_d is None or e_d is None:
            return "기간(시작일 ~ 종료일)을 선택하세요.", MUTED, False
        if s_d > e_d:
            return "종료일이 시작일보다 빠릅니다.", RED, False
        if state["done_file"]:
            return f"✓ 완료 — {state['done_file']}", GREEN, True
        return "준비됨", MUTED, True

    def refresh():
        for key, placeholder in (
            ("source", "파일을 선택하세요"),
            ("output", "파일을 선택하세요"),
        ):
            value = state[key]
            file_fields[key].config(text=value or placeholder, fg=TEXT if value else PLACEHOLDER)

        start_text.config(text=fmt_pair(state["start"]) or "날짜 선택", fg=TEXT if state["start"] else PLACEHOLDER)
        end_text.config(text=fmt_pair(state["end"]) or "날짜 선택", fg=TEXT if state["end"] else PLACEHOLDER)

        s_d, e_d = to_date(state["start"]), to_date(state["end"])
        if s_d is not None and e_d is not None and s_d <= e_d:
            days_label.config(text=f"{(e_d - s_d).days + 1}일간")
        else:
            days_label.config(text="")

        msg, color, can_run = validation_status()
        if state["running"]:
            status_label.config(text="생성 중…", fg=MUTED)
            if not progress.winfo_ismapped():
                progress.pack(side="left", padx=(10, 0))
                progress.start(12)
            can_run = False
        else:
            status_label.config(text=msg, fg=color)
            if progress.winfo_ismapped():
                progress.stop()
                progress.pack_forget()

        if can_run:
            run_btn.config(bg=ACCENT, fg="#ffffff", highlightbackground=ACCENT_DARK)
        else:
            run_btn.config(bg=DISABLED_BG, fg=DISABLED_FG, highlightbackground="#d4d4d4")
        run_btn.can_run = can_run

    def build_result():
        return {
            "mode": "cost_report",
            "use_template": False,
            "file_path": state["source"],
            "template_file": "",
            "output_file": state["output"],
            "year": str(current_year()),
            "start_month": str(state["start"][0]),
            "start_day": str(state["start"][1]),
            "end_month": str(state["end"][0]),
            "end_day": str(state["end"][1]),
        }

    def on_execute(_event=None):
        if state["running"] or not getattr(run_btn, "can_run", False):
            return
        close_picker()
        result = build_result()

        if on_run is None:
            legacy_result.update(result)
            root.quit()
            return

        state["running"] = True
        state["done_file"] = ""
        refresh()
        holder = {}

        def work():
            try:
                holder["lines"] = on_run(result)
            except Exception as exc:
                holder["error"] = exc

        worker = threading.Thread(target=work, daemon=True)
        worker.start()

        def poll():
            if worker.is_alive():
                root.after(100, poll)
                return
            state["running"] = False
            if "error" in holder:
                refresh()
                messagebox.showerror("원가 결과물 생성 실패", str(holder["error"]))
                return
            state["done_file"] = Path(result["output_file"]).name
            refresh()
            lines = holder.get("lines") or []
            messagebox.showinfo("원가 결과물 생성 완료", "\n".join(lines))

        root.after(100, poll)

    run_btn.bind("<Button-1>", on_execute)
    run_btn.bind("<Enter>", lambda _e: getattr(run_btn, "can_run", False) and not state["running"] and run_btn.config(bg=ACCENT_DARK))
    run_btn.bind("<Leave>", lambda _e: refresh())

    refresh()
    root.mainloop()
    try:
        root.destroy()
    except tk.TclError:
        pass  # 타이틀바 X 버튼으로 닫으면 root가 이미 파괴된 상태다.

    return legacy_result
