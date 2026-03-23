import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
from pathlib import Path


def run_ui():
    inventory_placeholder = "재고관리대장 파일을 선택하세요"
    template_placeholder = "결과물 템플릿 파일을 선택하세요"
    template_off_text = "템플릿 사용 안 함 (신규 생성 모드)"
    output_placeholder = "저장할 파일 경로를 선택하세요 (미선택시 자동 생성)"

    result = {}
    selected_template_file = {"path": ""}

    def select_inventory_file():
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            inventory_file_label.config(text=file_path)
            if use_template_var.get() and not selected_template_file["path"]:
                p = Path(file_path)
                candidate = p.parent / "결과물.xlsx"
                if candidate.exists():
                    selected_template_file["path"] = str(candidate)
            refresh_template_ui()

    def select_template_file():
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            selected_template_file["path"] = file_path
            refresh_template_ui()

    def refresh_template_ui():
        if use_template_var.get():
            template_file_button.config(state=tk.NORMAL)
            template_file_label.config(text=selected_template_file["path"] or template_placeholder)
        else:
            template_file_button.config(state=tk.DISABLED)
            template_file_label.config(text=template_off_text)

    def on_template_toggle():
        refresh_template_ui()

    def select_output_file():
        initial = "결과물_자동생성.xlsx"
        if use_template_var.get() and selected_template_file["path"]:
            initial = f"{Path(selected_template_file['path']).stem}_자동생성.xlsx"
        elif inventory_file_label.cget("text") != inventory_placeholder:
            initial = f"{Path(inventory_file_label.cget('text')).stem}_원가결과_자동생성.xlsx"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=initial,
        )
        if file_path:
            output_file_label.config(text=file_path)

    def on_run():
        inventory_file = inventory_file_label.cget("text")
        use_template = bool(use_template_var.get())
        template_file = selected_template_file["path"] if use_template else ""
        output_file = output_file_label.cget("text")
        year = year_var.get()
        start_month = start_month_var.get()
        start_day = start_day_var.get()
        end_month = end_month_var.get()
        end_day = end_day_var.get()

        if not inventory_file or inventory_file == inventory_placeholder:
            messagebox.showerror("오류", "재고관리대장 엑셀 파일을 선택하세요.")
            return

        if use_template and not template_file:
            messagebox.showerror("오류", "결과물 템플릿 엑셀 파일을 선택하세요.")
            return

        if not output_file or output_file == output_placeholder:
            base = Path(template_file) if use_template else Path(inventory_file)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = str(base.parent / f"{base.stem}_자동생성_{ts}.xlsx")
        if use_template and Path(template_file).resolve(strict=False) == Path(output_file).resolve(strict=False):
            base = Path(template_file)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = str(base.parent / f"{base.stem}_자동생성_{ts}.xlsx")

        result["mode"] = "cost_report"
        result["use_template"] = use_template
        result["file_path"] = inventory_file
        result["template_file"] = template_file
        result["output_file"] = output_file
        result["year"] = year
        result["start_month"] = start_month
        result["start_day"] = start_day
        result["end_month"] = end_month
        result["end_day"] = end_day

        root.quit()

    root = tk.Tk()
    root.title("원가 결과물 생성")
    root.geometry("760x400")

    tk.Button(root, text="재고관리대장 파일 선택", command=select_inventory_file).pack(pady=5)
    inventory_file_label = tk.Label(root, text=inventory_placeholder, anchor="w")
    inventory_file_label.pack(fill="x", padx=12)

    use_template_var = tk.BooleanVar(value=True)
    tk.Checkbutton(root, text="결과물 템플릿 사용", variable=use_template_var, command=on_template_toggle).pack(pady=3)

    template_file_button = tk.Button(root, text="결과물 템플릿 파일 선택", command=select_template_file)
    template_file_button.pack(pady=5)
    template_file_label = tk.Label(root, text=template_placeholder, anchor="w")
    template_file_label.pack(fill="x", padx=12)

    tk.Button(root, text="저장 경로 선택(선택사항)", command=select_output_file).pack(pady=5)
    output_file_label = tk.Label(root, text=output_placeholder, anchor="w")
    output_file_label.pack(fill="x", padx=12)

    year_var = tk.StringVar(value="2025")
    tk.Label(root, text="연도 선택").pack()
    tk.Entry(root, textvariable=year_var, width=10).pack()

    tk.Label(root, text="기간 선택 (월-일 ~ 월-일)").pack()
    frame = tk.Frame(root)
    frame.pack()

    start_month_var = tk.StringVar(value="07")
    start_day_var = tk.StringVar(value="24")

    end_month_var = tk.StringVar(value="08")
    end_day_var = tk.StringVar(value="25")

    tk.Entry(frame, textvariable=start_month_var, width=3).pack(side=tk.LEFT)
    tk.Label(frame, text="-").pack(side=tk.LEFT)
    tk.Entry(frame, textvariable=start_day_var, width=3).pack(side=tk.LEFT)
    tk.Label(frame, text=" ~ ").pack(side=tk.LEFT)
    tk.Entry(frame, textvariable=end_month_var, width=3).pack(side=tk.LEFT)
    tk.Label(frame, text="-").pack(side=tk.LEFT)
    tk.Entry(frame, textvariable=end_day_var, width=3).pack(side=tk.LEFT)

    tk.Button(root, text="실행", command=on_run).pack(pady=10)
    refresh_template_ui()

    root.mainloop()
    root.destroy()

    return result
