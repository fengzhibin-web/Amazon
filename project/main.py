import math
import os
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd

APP_TITLE = "Excel 商品数据提取工具"
WINDOW_SIZE = "760x520"


class ExcelExtractorApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.resizable(False, False)

        self.input_path_var = tk.StringVar()
        self.output_path_var = tk.StringVar()
        self.gemini_output_dir_var = tk.StringVar()
        self.batch_size_var = tk.StringVar(value="10")

        self._build_ui()

    def _build_ui(self):
        pad_x = 12
        pad_y = 8

        frame = tk.Frame(self.root)
        frame.pack(fill="both", expand=True, padx=pad_x, pady=pad_y)

        tk.Label(
            frame,
            text="功能1：按 id 去重并导出 Excel  |  功能2：生成 Gemini 批量输入 TXT",
            anchor="w",
            font=("Microsoft YaHei", 10, "bold"),
        ).pack(fill="x", pady=(0, 10))

        input_row = tk.Frame(frame)
        input_row.pack(fill="x", pady=(0, 8))
        tk.Label(input_row, text="输入 Excel:", width=14, anchor="w").pack(side="left")
        tk.Entry(input_row, textvariable=self.input_path_var).pack(side="left", fill="x", expand=True)
        tk.Button(input_row, text="选择", width=10, command=self.choose_input_file).pack(side="left", padx=(8, 0))

        output_row = tk.Frame(frame)
        output_row.pack(fill="x", pady=(0, 8))
        tk.Label(output_row, text="输出 Excel:", width=14, anchor="w").pack(side="left")
        tk.Entry(output_row, textvariable=self.output_path_var).pack(side="left", fill="x", expand=True)
        tk.Button(output_row, text="选择", width=10, command=self.choose_output_file).pack(side="left", padx=(8, 0))

        gemini_output_row = tk.Frame(frame)
        gemini_output_row.pack(fill="x", pady=(0, 8))
        tk.Label(gemini_output_row, text="Gemini 输出目录:", width=14, anchor="w").pack(side="left")
        tk.Entry(gemini_output_row, textvariable=self.gemini_output_dir_var).pack(side="left", fill="x", expand=True)
        tk.Button(gemini_output_row, text="选择", width=10, command=self.choose_gemini_output_dir).pack(
            side="left", padx=(8, 0)
        )

        batch_row = tk.Frame(frame)
        batch_row.pack(fill="x", pady=(0, 10))
        tk.Label(batch_row, text="批量大小:", width=14, anchor="w").pack(side="left")
        tk.Entry(batch_row, textvariable=self.batch_size_var, width=12).pack(side="left")
        tk.Label(batch_row, text="（默认 10，每 N 条产品生成一个 TXT）", fg="#475569").pack(side="left", padx=(10, 0))

        action_row = tk.Frame(frame)
        action_row.pack(fill="x", pady=(2, 10))
        tk.Button(
            action_row,
            text="开始提取 Excel",
            width=18,
            height=2,
            command=self.run_extract,
            bg="#0ea5e9",
            fg="white",
            activebackground="#0284c7",
            activeforeground="white",
            relief="flat",
        ).pack(side="left")

        tk.Button(
            action_row,
            text="生成 Gemini 批量输入",
            width=20,
            height=2,
            command=self.generate_gemini_batches,
            bg="#16a34a",
            fg="white",
            activebackground="#15803d",
            activeforeground="white",
            relief="flat",
        ).pack(side="left", padx=(10, 0))

        help_text = (
            "Excel 列名要求：\n"
            "1) 必须存在 id 列（不区分大小写）\n"
            "2) 必须存在 英语-标题 列\n"
            "3) 必须存在 英语-描述 列\n\n"
            "Gemini 文本格式：\n"
            "[PRODUCT] / id / 原始标题 / 原始描述 / [/PRODUCT]"
        )
        tk.Label(frame, text=help_text, justify="left", fg="#334155").pack(fill="x", pady=(0, 8))

        self.log_text = tk.Text(frame, height=12, wrap="word", state="disabled")
        self.log_text.pack(fill="both", expand=True)

    def choose_input_file(self):
        file_path = filedialog.askopenfilename(
            title="选择输入 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if file_path:
            self.input_path_var.set(file_path)
            if not self.output_path_var.get():
                self.output_path_var.set(self._default_output_path(file_path))
            if not self.gemini_output_dir_var.get():
                self.gemini_output_dir_var.set(self._default_gemini_output_dir(file_path))

    def choose_output_file(self):
        file_path = filedialog.asksaveasfilename(
            title="选择输出 Excel 文件",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialfile="extracted_result.xlsx",
        )
        if file_path:
            self.output_path_var.set(file_path)

    def choose_gemini_output_dir(self):
        dir_path = filedialog.askdirectory(title="选择 Gemini TXT 输出目录")
        if dir_path:
            self.gemini_output_dir_var.set(dir_path)

    @staticmethod
    def _default_output_path(input_path: str) -> str:
        base_dir = os.path.dirname(input_path)
        name, _ = os.path.splitext(os.path.basename(input_path))
        return os.path.join(base_dir, f"{name}_extracted.xlsx")

    @staticmethod
    def _default_gemini_output_dir(input_path: str) -> str:
        base_dir = os.path.dirname(input_path)
        return os.path.join(base_dir, "gemini_batches")

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

    def _load_required_columns(self, input_path: str) -> pd.DataFrame:
        df = pd.read_excel(input_path)
        if df.empty:
            raise ValueError("输入 Excel 为空，没有可处理的数据。")

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
            raise ValueError(f"缺少必要列: {', '.join(missing)}")

        work_df = df[[id_col, title_col, desc_col]].copy()
        work_df.columns = ["id", "英语-标题", "英语-描述"]
        work_df["id"] = work_df["id"].astype(str).str.strip()
        work_df = work_df[work_df["id"] != ""].copy()

        # 标题和描述为空时写空字符串；描述允许为空
        work_df["英语-标题"] = work_df["英语-标题"].fillna("").astype(str)
        work_df["英语-描述"] = work_df["英语-描述"].fillna("").astype(str)
        return work_df

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
            work_df = self._load_required_columns(input_path)

            before_count = len(work_df)
            dedup_df = work_df.drop_duplicates(subset=["id"], keep="first")
            after_count = len(dedup_df)

            output_dir = os.path.dirname(output_path)
            if output_dir:
                os.makedirs(output_dir, exist_ok=True)

            dedup_df.to_excel(output_path, index=False)

            self._append_log(f"原始记录数: {before_count}")
            self._append_log(f"去重后记录数: {after_count}")
            self._append_log(f"导出完成: {output_path}")
            messagebox.showinfo("完成", f"Excel 提取完成！\n输出文件：\n{output_path}")

        except Exception as exc:
            self._append_log("发生错误：")
            self._append_log(str(exc))
            self._append_log(traceback.format_exc())
            messagebox.showerror("错误", f"处理失败：{exc}")

    def generate_gemini_batches(self):
        input_path = self.input_path_var.get().strip()
        output_dir = self.gemini_output_dir_var.get().strip()
        batch_size_str = self.batch_size_var.get().strip()

        if not input_path:
            messagebox.showwarning("提示", "请先选择输入 Excel 文件。")
            return
        if not os.path.exists(input_path):
            messagebox.showerror("错误", "输入文件不存在，请重新选择。")
            return
        if not output_dir:
            output_dir = self._default_gemini_output_dir(input_path)
            self.gemini_output_dir_var.set(output_dir)

        try:
            batch_size = int(batch_size_str or "10")
            if batch_size <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("错误", "批量大小必须是大于 0 的整数。")
            return

        try:
            self._append_log(f"读取文件用于 Gemini 批量: {input_path}")
            work_df = self._load_required_columns(input_path)
            dedup_df = work_df.drop_duplicates(subset=["id"], keep="first").reset_index(drop=True)

            if dedup_df.empty:
                messagebox.showwarning("提示", "去重后没有可生成的产品数据。")
                return

            os.makedirs(output_dir, exist_ok=True)

            total = len(dedup_df)
            file_count = int(math.ceil(total / batch_size))

            for file_index in range(file_count):
                start = file_index * batch_size
                end = min(start + batch_size, total)
                chunk = dedup_df.iloc[start:end]

                filename = f"gemini_batch_{file_index + 1:03d}.txt"
                file_path = os.path.join(output_dir, filename)

                blocks = []
                for _, row in chunk.iterrows():
                    product_id = str(row["id"]).strip()
                    title = str(row["英语-标题"]) if pd.notna(row["英语-标题"]) else ""
                    desc = str(row["英语-描述"]) if pd.notna(row["英语-描述"]) else ""

                    block = (
                        "[PRODUCT]\n"
                        f"id: {product_id}\n"
                        f"原始标题: {title}\n"
                        f"原始描述: {desc}\n"
                        "[/PRODUCT]"
                    )
                    blocks.append(block)

                content = "\n\n".join(blocks)
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write(content)

            self._append_log(f"Gemini 批量生成完成：共 {file_count} 个 txt，输出目录: {output_dir}")
            messagebox.showinfo("完成", f"Gemini 批量输入生成完成！\n共生成 {file_count} 个 txt 文件。")

        except Exception as exc:
            self._append_log("Gemini 生成发生错误：")
            self._append_log(str(exc))
            self._append_log(traceback.format_exc())
            messagebox.showerror("错误", f"生成失败：{exc}")


def main():
    root = tk.Tk()
    app = ExcelExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
