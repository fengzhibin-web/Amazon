import os
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd

APP_TITLE = "Excel 商品数据提取工具"
WINDOW_SIZE = "620x360"


class ExcelExtractorApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.resizable(False, False)

        self.input_path_var = tk.StringVar()
        self.output_path_var = tk.StringVar()

        self._build_ui()

    def _build_ui(self):
        pad_x = 12
        pad_y = 8

        frame = tk.Frame(self.root)
        frame.pack(fill="both", expand=True, padx=pad_x, pady=pad_y)

        tk.Label(
            frame,
            text="功能：按 id 去重，提取“英语-标题”“英语-描述”并导出新 Excel",
            anchor="w",
            font=("Microsoft YaHei", 10, "bold"),
        ).pack(fill="x", pady=(0, 10))

        input_row = tk.Frame(frame)
        input_row.pack(fill="x", pady=(0, 8))
        tk.Label(input_row, text="输入 Excel:", width=12, anchor="w").pack(side="left")
        tk.Entry(input_row, textvariable=self.input_path_var).pack(side="left", fill="x", expand=True)
        tk.Button(input_row, text="选择", width=10, command=self.choose_input_file).pack(side="left", padx=(8, 0))

        output_row = tk.Frame(frame)
        output_row.pack(fill="x", pady=(0, 8))
        tk.Label(output_row, text="输出 Excel:", width=12, anchor="w").pack(side="left")
        tk.Entry(output_row, textvariable=self.output_path_var).pack(side="left", fill="x", expand=True)
        tk.Button(output_row, text="选择", width=10, command=self.choose_output_file).pack(side="left", padx=(8, 0))

        action_row = tk.Frame(frame)
        action_row.pack(fill="x", pady=(4, 10))
        tk.Button(
            action_row,
            text="开始提取",
            width=16,
            height=2,
            command=self.run_extract,
            bg="#0ea5e9",
            fg="white",
            activebackground="#0284c7",
            activeforeground="white",
            relief="flat",
        ).pack(side="left")

        help_text = (
            "列名要求：\n"
            "1) 必须存在 id 列（不区分大小写）\n"
            "2) 必须存在 英语-标题 列\n"
            "3) 必须存在 英语-描述 列\n\n"
            "输出列：id、英语-标题、英语-描述"
        )
        tk.Label(frame, text=help_text, justify="left", fg="#334155").pack(fill="x", pady=(4, 8))

        self.log_text = tk.Text(frame, height=8, wrap="word", state="disabled")
        self.log_text.pack(fill="both", expand=True)

    def choose_input_file(self):
        file_path = filedialog.askopenfilename(
            title="选择输入 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if file_path:
            self.input_path_var.set(file_path)
            if not self.output_path_var.get():
                default_output = self._default_output_path(file_path)
                self.output_path_var.set(default_output)

    def choose_output_file(self):
        file_path = filedialog.asksaveasfilename(
            title="选择输出 Excel 文件",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialfile="extracted_result.xlsx",
        )
        if file_path:
            self.output_path_var.set(file_path)

    @staticmethod
    def _default_output_path(input_path: str) -> str:
        base_dir = os.path.dirname(input_path)
        name, _ = os.path.splitext(os.path.basename(input_path))
        return os.path.join(base_dir, f"{name}_extracted.xlsx")

    def _append_log(self, msg: str):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    @staticmethod
    def _find_column(columns, target_name: str, case_insensitive: bool = False):
        if case_insensitive:
            lookup = {str(c).strip().lower(): c for c in columns}
            return lookup.get(target_name.strip().lower())
        lookup = {str(c).strip(): c for c in columns}
        return lookup.get(target_name.strip())

    def run_extract(self):
        input_path = self.input_path_var.get().strip()
        output_path = self.output_path_var.get().strip()

        if not input_path:
            messagebox.showwarning("提示", "请先选择输入 Excel 文件。")
            return
        if not os.path.exists(input_path):
            messagebox.showerror("错误", "输入文件不存在，请重新选择。")
            return

        if not output_path:
            output_path = self._default_output_path(input_path)
            self.output_path_var.set(output_path)

        try:
            self._append_log(f"读取文件: {input_path}")
            df = pd.read_excel(input_path)
            if df.empty:
                messagebox.showwarning("提示", "输入 Excel 为空，没有可处理的数据。")
                return

            id_col = self._find_column(df.columns, "id", case_insensitive=True)
            title_col = self._find_column(df.columns, "英语-标题")
            desc_col = self._find_column(df.columns, "英语-描述")

            missing = []
            if id_col is None:
                missing.append("id")
            if title_col is None:
                missing.append("英语-标题")
            if desc_col is None:
                missing.append("英语-描述")

            if missing:
                messagebox.showerror("列缺失", f"缺少必要列: {', '.join(missing)}")
                self._append_log("列检查失败，处理终止。")
                return

            work_df = df[[id_col, title_col, desc_col]].copy()
            work_df.columns = ["id", "英语-标题", "英语-描述"]

            before_count = len(work_df)
            work_df["id"] = work_df["id"].astype(str).str.strip()
            work_df = work_df[work_df["id"] != ""]
            work_df = work_df.drop_duplicates(subset=["id"], keep="first")
            after_count = len(work_df)

            output_dir = os.path.dirname(output_path)
            if output_dir:
                os.makedirs(output_dir, exist_ok=True)

            work_df.to_excel(output_path, index=False)

            self._append_log(f"原始记录数: {before_count}")
            self._append_log(f"去重后记录数: {after_count}")
            self._append_log(f"导出完成: {output_path}")
            messagebox.showinfo("完成", f"处理完成！\n输出文件：\n{output_path}")

        except Exception as exc:
            self._append_log("发生错误：")
            self._append_log(str(exc))
            self._append_log(traceback.format_exc())
            messagebox.showerror("错误", f"处理失败：{exc}")


def main():
    root = tk.Tk()
    app = ExcelExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
