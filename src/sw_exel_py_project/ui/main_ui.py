import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
from pathlib import Path


def run_ui():
    inventory_placeholder = "재고관리대장 파일을 선택하세요"
    template_placeholder = "결과물 템플릿 파일을 선택하세요"
    output_placeholder = "저장할 파일 경로를 선택하세요"

    result = {}
    selected_template_file = {"path": ""}

    def select_inventory_file():
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            inventory_file_label.config(text=file_path)
            if not selected_template_file["path"]:
                p = Path(file_path)
                candidate = p.parent / "결과물.xlsx"
                if candidate.exists():
                    selected_template_file["path"] = str(candidate)
            template_file_label.config(text=selected_template_file["path"] or template_placeholder)

    def select_template_file():
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            selected_template_file["path"] = file_path
            template_file_label.config(text=file_path)

    def select_output_file():
        initial = "결과물_자동생성.xlsx"
        if selected_template_file["path"]:
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
        template_file = selected_template_file["path"]
        output_file = output_file_label.cget("text")

        if not inventory_file or inventory_file == inventory_placeholder:
            messagebox.showerror("오류", "재고관리대장 엑셀 파일을 선택하세요.")
            return

        if not template_file:
            messagebox.showerror("오류", "결과물 템플릿 엑셀 파일을 선택하세요.")
            return

        if not output_file or output_file == output_placeholder:
            messagebox.showerror("오류", "저장할 파일 경로를 선택하세요.")
            return

        if Path(template_file).resolve(strict=False) == Path(output_file).resolve(strict=False):
            messagebox.showerror("오류", "저장 경로가 템플릿 파일과 같습니다. 다른 경로를 선택하세요.")
            return

        try:
            year = int(year_var.get())
            start = datetime(year, int(start_month_var.get()), int(start_day_var.get()))
            end = datetime(year, int(end_month_var.get()), int(end_day_var.get()))
        except ValueError:
            messagebox.showerror("오류", "연도와 기간(월-일)을 올바르게 입력하세요.")
            return

        if start > end:
            messagebox.showerror("오류", "시작일이 종료일보다 늦습니다.")
            return

        result["mode"] = "cost_report"
        result["use_template"] = True
        result["file_path"] = inventory_file
        result["template_file"] = template_file
        result["output_file"] = output_file
        result["year"] = str(year)
        result["start_month"] = start_month_var.get()
        result["start_day"] = start_day_var.get()
        result["end_month"] = end_month_var.get()
        result["end_day"] = end_day_var.get()

        root.quit()

    root = tk.Tk()
    root.title("원가 결과물 생성")
    root.geometry("760x400")

    tk.Button(root, text="재고관리대장 파일 선택", command=select_inventory_file).pack(pady=5)
    inventory_file_label = tk.Label(root, text=inventory_placeholder, anchor="w")
    inventory_file_label.pack(fill="x", padx=12)

    tk.Button(root, text="결과물 템플릿 파일 선택", command=select_template_file).pack(pady=5)
    template_file_label = tk.Label(root, text=template_placeholder, anchor="w")
    template_file_label.pack(fill="x", padx=12)

    tk.Button(root, text="저장 경로 선택", command=select_output_file).pack(pady=5)
    output_file_label = tk.Label(root, text=output_placeholder, anchor="w")
    output_file_label.pack(fill="x", padx=12)

    year_var = tk.StringVar(value=str(datetime.now().year))
    tk.Label(root, text="연도 선택").pack()
    tk.Entry(root, textvariable=year_var, width=10).pack()

    tk.Label(root, text="기간 선택 (월-일 ~ 월-일)").pack()
    frame = tk.Frame(root)
    frame.pack()

    start_month_var = tk.StringVar(value="")
    start_day_var = tk.StringVar(value="")

    end_month_var = tk.StringVar(value="")
    end_day_var = tk.StringVar(value="")

    tk.Entry(frame, textvariable=start_month_var, width=3).pack(side=tk.LEFT)
    tk.Label(frame, text="-").pack(side=tk.LEFT)
    tk.Entry(frame, textvariable=start_day_var, width=3).pack(side=tk.LEFT)
    tk.Label(frame, text=" ~ ").pack(side=tk.LEFT)
    tk.Entry(frame, textvariable=end_month_var, width=3).pack(side=tk.LEFT)
    tk.Label(frame, text="-").pack(side=tk.LEFT)
    tk.Entry(frame, textvariable=end_day_var, width=3).pack(side=tk.LEFT)

    tk.Button(root, text="실행", command=on_run).pack(pady=10)

    root.mainloop()
    root.destroy()

    return result
