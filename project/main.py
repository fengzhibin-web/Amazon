import io
import math
import os
import re
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill

APP_TITLE = "Excel 商品数据提取工具"
WINDOW_SIZE = "780x560"

INVISIBLE_CHAR_PATTERN = re.compile(r"[\u200b\u200c\u200d\ufeff]")


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
            text="功能1：Excel去重导出 | 功能2：Gemini批量TXT | 功能3：Gemini TSV转纵向Excel",
            anchor="w",
            font=("Microsoft YaHei", 10, "bold"),
        ).pack(fill="x", pady=(0, 10))

        input_row = tk.Frame(frame)
        input_row.pack(fill="x", pady=(0, 8))
        tk.Label(input_row, text="输入 Excel:", width=16, anchor="w").pack(side="left")
        tk.Entry(input_row, textvariable=self.input_path_var).pack(side="left", fill="x", expand=True)
        tk.Button(input_row, text="选择", width=10, command=self.choose_input_file).pack(side="left", padx=(8, 0))

        output_row = tk.Frame(frame)
        output_row.pack(fill="x", pady=(0, 8))
        tk.Label(output_row, text="输出 Excel:", width=16, anchor="w").pack(side="left")
        tk.Entry(output_row, textvariable=self.output_path_var).pack(side="left", fill="x", expand=True)
        tk.Button(output_row, text="选择", width=10, command=self.choose_output_file).pack(side="left", padx=(8, 0))

        gemini_output_row = tk.Frame(frame)
        gemini_output_row.pack(fill="x", pady=(0, 8))
        tk.Label(gemini_output_row, text="Gemini TXT 输出目录:", width=16, anchor="w").pack(side="left")
        tk.Entry(gemini_output_row, textvariable=self.gemini_output_dir_var).pack(side="left", fill="x", expand=True)
        tk.Button(gemini_output_row, text="选择", width=10, command=self.choose_gemini_output_dir).pack(
            side="left", padx=(8, 0)
        )

        batch_row = tk.Frame(frame)
        batch_row.pack(fill="x", pady=(0, 10))
        tk.Label(batch_row, text="批量大小:", width=16, anchor="w").pack(side="left")
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
            relief="flat",
        ).pack(side="left", padx=(10, 0))

        tk.Button(
            action_row,
            text="Gemini TSV转纵向表",
            width=18,
            height=2,
            command=self.run_convert_gemini_vertical,
            bg="#9333ea",
            fg="white",
            relief="flat",
        ).pack(side="left", padx=(10, 0))

        help_text = (
            "Excel 列名要求：id（不区分大小写）、英语-标题、英语-描述\n"
            "Gemini 批量TXT：每个文件顶部新增 IDS: id1,id2,id3...，其后空一行再写 [PRODUCT]\n"
            "Gemini 结果转换：支持 .txt/.tsv/.xlsx 输入，输出 gemini_result_vertical.xlsx"
        )
        tk.Label(frame, text=help_text, justify="left", fg="#334155").pack(fill="x", pady=(0, 8))

        self.log_text = tk.Text(frame, height=14, wrap="word", state="disabled")
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

    @staticmethod
    def _normalize_col_name(name: str) -> str:
        text = str(name)
        text = INVISIBLE_CHAR_PATTERN.sub("", text)
        return text.strip()

    @staticmethod
    def clean_gemini_code_block(content: str) -> str:
        text = content.replace("\r\n", "\n").replace("\r", "\n").strip()
        if text.startswith("```"):
            lines = text.split("\n")
            if lines and lines[0].strip().startswith("```"):
                lines = lines[1:]
            if lines and lines[-1].strip() == "```":
                lines = lines[:-1]
            text = "\n".join(lines)
        return text.strip("\n")

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

    def generate_gemini_batch_txt(self, dedup_df: pd.DataFrame, output_dir: str, batch_size: int) -> int:
        os.makedirs(output_dir, exist_ok=True)

        total = len(dedup_df)
        file_count = int(math.ceil(total / batch_size))

        for file_index in range(file_count):
            start = file_index * batch_size
            end = min(start + batch_size, total)
            chunk = dedup_df.iloc[start:end]

            filename = f"gemini_batch_{file_index + 1:03d}.txt"
            file_path = os.path.join(output_dir, filename)

            ids_line = "IDS: " + ",".join([str(v).strip() for v in chunk["id"].tolist()])

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

            content = ids_line + "\n\n" + "\n\n".join(blocks)
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(content)

        return file_count

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

            file_count = self.generate_gemini_batch_txt(dedup_df, output_dir, batch_size)
            self._append_log(f"Gemini 批量生成完成：共 {file_count} 个 txt，输出目录: {output_dir}")
            messagebox.showinfo("完成", f"Gemini 批量输入生成完成！\n共生成 {file_count} 个 txt 文件。")

        except Exception as exc:
            self._append_log("Gemini 生成发生错误：")
            self._append_log(str(exc))
            self._append_log(traceback.format_exc())
            messagebox.showerror("错误", f"生成失败：{exc}")

    def read_gemini_result_file(self, input_path: str) -> pd.DataFrame:
        ext = os.path.splitext(input_path)[1].lower()
        if ext in {".txt", ".tsv"}:
            with open(input_path, "r", encoding="utf-8") as f:
                raw = f.read()
            clean = self.clean_gemini_code_block(raw)
            df = pd.read_csv(io.StringIO(clean), sep="\t", dtype=str, keep_default_na=False)
        elif ext in {".xlsx", ".xls"}:
            df = pd.read_excel(input_path, dtype=str)
        else:
            raise ValueError("仅支持 .txt / .tsv / .xlsx / .xls 文件。")

        if df.empty:
            raise ValueError("输入文件为空，没有可转换的数据。")

        df.columns = [self._normalize_col_name(col) for col in df.columns]
        df = df.fillna("")
        return df

    @staticmethod
    def merge_numbered_fields(row: pd.Series, keys: list[str]) -> str:
        items = []
        for key in keys:
            value = str(row.get(key, "")).strip()
            if value:
                items.append(value)
        return "\n".join([f"{idx}. {val}" for idx, val in enumerate(items, start=1)])

    def convert_gemini_result_to_vertical_excel(self, input_path: str, output_dir: str) -> tuple[str, int]:
        df = self.read_gemini_result_file(input_path)

        required_map = {
            "id": "产品ID",
            "原始标题": "原始标题",
            "原始描述": "原始描述",
            "违禁词检验": "违禁词检验",
            "品牌识别": "品牌识别",
            "AI_通用标题_EN": "通用标题 EN",
            "AI_通用标题_CN": "通用标题 CN",
            "AI_关键词_EN": "关键词 EN",
            "AI_关键词_CN": "关键词 CN",
        }

        missing_fields = [k for k in required_map if k not in df.columns]
        if missing_fields:
            self._append_log(f"提示：输入缺少字段，将按空值处理: {', '.join(missing_fields)}")

        selling_en_keys = [f"AI_卖点_EN_{i}" for i in range(1, 6)]
        selling_cn_keys = [f"AI_卖点_CN_{i}" for i in range(1, 6)]
        spec_en_keys = [f"AI_规格描述_EN_{i}" for i in range(1, 6)]
        spec_cn_keys = [f"AI_规格描述_CN_{i}" for i in range(1, 6)]

        for key in selling_en_keys + selling_cn_keys + spec_en_keys + spec_cn_keys:
            if key not in df.columns:
                self._append_log(f"提示：缺少字段 {key}，已按空值处理")

        records = []
        for _, row in df.iterrows():
            field_rows = [
                ("产品ID", str(row.get("id", "")).strip()),
                ("AI状态", "成功"),
                ("原始标题", str(row.get("原始标题", "")).strip()),
                ("原始描述", str(row.get("原始描述", "")).strip()),
                ("违禁词检验", str(row.get("违禁词检验", "")).strip()),
                ("品牌识别", str(row.get("品牌识别", "")).strip()),
                ("通用标题 EN", str(row.get("AI_通用标题_EN", "")).strip()),
                ("通用标题 CN", str(row.get("AI_通用标题_CN", "")).strip()),
                ("关键词 EN", str(row.get("AI_关键词_EN", "")).strip()),
                ("关键词 CN", str(row.get("AI_关键词_CN", "")).strip()),
                ("卖点 EN", self.merge_numbered_fields(row, selling_en_keys)),
                ("卖点 CN", self.merge_numbered_fields(row, selling_cn_keys)),
                ("规格描述 EN", self.merge_numbered_fields(row, spec_en_keys)),
                ("规格描述 CN", self.merge_numbered_fields(row, spec_cn_keys)),
            ]

            for field, value in field_rows:
                records.append({"字段": field, "内容": value})
            records.append({"字段": "", "内容": ""})

        out_df = pd.DataFrame(records)
        os.makedirs(output_dir, exist_ok=True)
        out_path = os.path.join(output_dir, "gemini_result_vertical.xlsx")

        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            out_df.to_excel(writer, index=False, sheet_name="Vertical")
            ws = writer.sheets["Vertical"]
            self.apply_vertical_excel_styles(ws)

        return out_path, len(df)

    def apply_vertical_excel_styles(self, ws):
        header_fill = PatternFill(fill_type="solid", fgColor="DCEBFF")
        product_id_fill = PatternFill(fill_type="solid", fgColor="C7DFFF")

        ws.freeze_panes = "A2"

        ws.cell(row=1, column=1).font = Font(bold=True)
        ws.cell(row=1, column=2).font = Font(bold=True)
        ws.cell(row=1, column=1).fill = header_fill
        ws.cell(row=1, column=2).fill = header_fill

        for row_idx in range(2, ws.max_row + 1):
            field_cell = ws.cell(row=row_idx, column=1)
            content_cell = ws.cell(row=row_idx, column=2)

            field_cell.font = Font(bold=True)
            field_cell.fill = header_fill

            if str(field_cell.value).strip() == "产品ID":
                field_cell.fill = product_id_fill
                content_cell.fill = product_id_fill
                content_cell.font = Font(bold=True)

            content_cell.alignment = Alignment(wrap_text=True, vertical="top")
            field_cell.alignment = Alignment(vertical="top")

            text = str(content_cell.value or "")
            lines = text.count("\n") + 1
            ws.row_dimensions[row_idx].height = max(20, min(120, lines * 18))

        ws.column_dimensions["A"].width = 20
        max_len = 30
        for row_idx in range(2, ws.max_row + 1):
            value = str(ws.cell(row=row_idx, column=2).value or "")
            longest = max([len(line) for line in value.split("\n")], default=0)
            max_len = max(max_len, min(longest + 4, 100))
        ws.column_dimensions["B"].width = max_len

    def run_convert_gemini_vertical(self):
        input_path = filedialog.askopenfilename(
            title="选择 Gemini 结果文件",
            filetypes=[
                ("Gemini 结果", "*.txt *.tsv *.xlsx *.xls"),
                ("文本文件", "*.txt *.tsv"),
                ("Excel 文件", "*.xlsx *.xls"),
                ("所有文件", "*.*"),
            ],
        )
        if not input_path:
            return

        output_dir = filedialog.askdirectory(title="选择纵向 Excel 输出目录")
        if not output_dir:
            return

        try:
            out_path, count = self.convert_gemini_result_to_vertical_excel(input_path, output_dir)
            self._append_log(f"转换完成，共处理 {count} 个产品，输出文件：{out_path}")
            messagebox.showinfo("完成", f"转换完成，共处理 {count} 个产品。\n输出文件：\n{out_path}")
        except Exception as exc:
            self._append_log("Gemini 纵向转换发生错误：")
            self._append_log(str(exc))
            self._append_log(traceback.format_exc())
            messagebox.showerror("错误", f"转换失败：{exc}")


def main():
    root = tk.Tk()
    ExcelExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
