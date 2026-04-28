import json
import math
import os
import re
import traceback
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    DND_AVAILABLE = True
except Exception:
    DND_AVAILABLE = False
    DND_FILES = None
    TkinterDnD = None

APP_TITLE = "Excel 商品数据提取工具"
WINDOW_SIZE = "1180x760"
INVISIBLE_CHAR_PATTERN = re.compile(r"[\u200b\u200c\u200d\ufeff]")

GEMINI_FIELD_ORDER = [
    "id",
    "原始标题",
    "原始描述",
    "违禁词检验",
    "品牌识别",
    "AI_通用标题_EN",
    "AI_通用标题_CN",
    "AI_关键词_EN",
    "AI_关键词_CN",
    "AI_卖点_EN_1",
    "AI_卖点_EN_2",
    "AI_卖点_EN_3",
    "AI_卖点_EN_4",
    "AI_卖点_EN_5",
    "AI_卖点_CN_1",
    "AI_卖点_CN_2",
    "AI_卖点_CN_3",
    "AI_卖点_CN_4",
    "AI_卖点_CN_5",
    "AI_规格描述_EN",
    "AI_规格描述_CN",
]

OLD_SPEC_EN_FIELDS = [f"AI_规格描述_EN_{i}" for i in range(1, 6)]
OLD_SPEC_CN_FIELDS = [f"AI_规格描述_CN_{i}" for i in range(1, 6)]

NUMBER_CLEAN_FIELDS = {
    "AI_卖点_EN_1",
    "AI_卖点_EN_2",
    "AI_卖点_EN_3",
    "AI_卖点_EN_4",
    "AI_卖点_EN_5",
    "AI_卖点_CN_1",
    "AI_卖点_CN_2",
    "AI_卖点_CN_3",
    "AI_卖点_CN_4",
    "AI_卖点_CN_5",
    "AI_规格描述_EN",
    "AI_规格描述_CN",
}



class ExcelExtractorApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry(WINDOW_SIZE)

        self.input_path_var = tk.StringVar()
        self.output_path_var = tk.StringVar()
        self.gemini_output_dir_var = tk.StringVar()
        self.batch_size_var = tk.StringVar(value="10")

        self.gemini_source_file_var = tk.StringVar()
        self.vertical_output_dir_var = tk.StringVar()
        self.vertical_filename_var = tk.StringVar()

        self.tasks = []
        self.task_counter = 0

        self._build_ui()

    def _build_ui(self):
        top_frame = tk.Frame(self.root)
        top_frame.pack(fill="both", expand=True, padx=12, pady=8)

        tk.Label(
            top_frame,
            text="Excel去重 | Gemini批量TXT | Gemini TSV转纵向表 | 任务列表",
            anchor="w",
            font=("Microsoft YaHei", 10, "bold"),
        ).pack(fill="x", pady=(0, 8))

        panes = tk.PanedWindow(top_frame, orient="horizontal", sashrelief="raised")
        panes.pack(fill="both", expand=True)

        left = tk.Frame(panes)
        right = tk.Frame(panes)
        panes.add(left, minsize=700)
        panes.add(right, minsize=420)

        self._build_left_area(left)
        self._build_task_area(right)

    def _build_left_area(self, parent):
        notebook = ttk.Notebook(parent)
        notebook.pack(fill="both", expand=True)

        tab_excel = tk.Frame(notebook)
        tab_batch = tk.Frame(notebook)
        tab_vertical = tk.Frame(notebook)

        notebook.add(tab_excel, text="Excel去重提取")
        notebook.add(tab_batch, text="生成Gemini批量输入")
        notebook.add(tab_vertical, text="Gemini TSV转纵向表")

        self._build_tab_excel(tab_excel)
        self._build_tab_batch(tab_batch)
        self._build_tab_vertical(tab_vertical)

        self.log_text = tk.Text(parent, height=10, wrap="word", state="disabled")
        self.log_text.pack(fill="both", expand=True, pady=(8, 0))

    def _build_tab_excel(self, parent):
        self._labeled_path_row(parent, "输入 Excel:", self.input_path_var, self.choose_input_file, drop_types={".xlsx", ".xls"})
        self._labeled_path_row(parent, "输出 Excel:", self.output_path_var, self.choose_output_file, save_mode=True)
        self._add_drop_zone(parent, "拖拽上传区（Excel去重提取）", self.input_path_var, {".xlsx", ".xls"})

        tk.Button(parent, text="开始提取 Excel", width=20, command=self.run_extract, bg="#0ea5e9", fg="white").pack(
            anchor="w", pady=8
        )

    def _build_tab_batch(self, parent):
        self._labeled_path_row(
            parent,
            "输入 Excel:",
            self.input_path_var,
            self.choose_input_file,
            drop_types={".xlsx", ".xls"},
        )
        self._labeled_path_row(
            parent,
            "Gemini TXT 输出目录:",
            self.gemini_output_dir_var,
            self.choose_gemini_output_dir,
            dir_mode=True,
        )
        self._add_drop_zone(parent, "拖拽上传区（Gemini批量输入）", self.input_path_var, {".xlsx", ".xls"})

        row = tk.Frame(parent)
        row.pack(fill="x", pady=4)
        tk.Label(row, text="批量大小:", width=20, anchor="w").pack(side="left")
        tk.Entry(row, textvariable=self.batch_size_var, width=12).pack(side="left")

        tk.Button(parent, text="生成 Gemini 批量输入", width=24, command=self.generate_gemini_batches, bg="#16a34a", fg="white").pack(
            anchor="w", pady=8
        )

    def _build_tab_vertical(self, parent):
        self._labeled_path_row(
            parent,
            "Gemini 输入文件:",
            self.gemini_source_file_var,
            self.choose_gemini_source_file,
            drop_types={".txt", ".tsv", ".xlsx", ".xls"},
        )
        self._add_drop_zone(parent, "拖拽上传区（Gemini TSV转纵向表）", self.gemini_source_file_var, {".txt", ".tsv", ".xlsx", ".xls"})

        row_out = tk.Frame(parent)
        row_out.pack(fill="x", pady=4)
        tk.Label(row_out, text="输出目录:", width=20, anchor="w").pack(side="left")
        tk.Entry(row_out, textvariable=self.vertical_output_dir_var).pack(side="left", fill="x", expand=True)
        tk.Button(row_out, text="选择", width=10, command=self.choose_vertical_output_dir).pack(side="left", padx=(8, 0))

        row_name = tk.Frame(parent)
        row_name.pack(fill="x", pady=4)
        tk.Label(row_name, text="输出文件名:", width=20, anchor="w").pack(side="left")
        tk.Entry(row_name, textvariable=self.vertical_filename_var).pack(side="left", fill="x", expand=True)

        tk.Label(parent, text="粘贴 Gemini TSV 文本（支持 ```tsv 包裹）:", anchor="w").pack(fill="x", pady=(8, 4))
        self.tsv_paste_text = tk.Text(parent, height=10, wrap="word")
        self.tsv_paste_text.pack(fill="both", expand=True)

        btn_row = tk.Frame(parent)
        btn_row.pack(fill="x", pady=8)
        tk.Button(btn_row, text="从文件转换", width=16, command=self.run_convert_gemini_vertical_from_file, bg="#9333ea", fg="white").pack(
            side="left"
        )
        tk.Button(btn_row, text="从粘贴文本转换", width=18, command=self.run_convert_gemini_vertical_from_paste, bg="#7e22ce", fg="white").pack(
            side="left", padx=(8, 0)
        )

    def _build_task_area(self, parent):
        tk.Label(parent, text="任务列表", font=("Microsoft YaHei", 10, "bold")).pack(anchor="w", pady=(0, 6))

        columns = ("task_name", "source", "status", "count", "output", "created", "finished", "remark")
        tree = ttk.Treeview(parent, columns=columns, show="headings", height=30)
        headers = {
            "task_name": "任务名",
            "source": "输入来源",
            "status": "状态",
            "count": "产品数量",
            "output": "输出文件",
            "created": "创建时间",
            "finished": "完成时间",
            "remark": "备注",
        }
        widths = {
            "task_name": 120,
            "source": 100,
            "status": 70,
            "count": 70,
            "output": 150,
            "created": 130,
            "finished": 130,
            "remark": 250,
        }
        for col in columns:
            tree.heading(col, text=headers[col])
            tree.column(col, width=widths[col], anchor="w")

        tree.pack(fill="both", expand=True)
        self.task_tree = tree

    def _labeled_path_row(self, parent, label, var, command, save_mode=False, dir_mode=False, drop_types=None):
        row = tk.Frame(parent)
        row.pack(fill="x", pady=4)
        tk.Label(row, text=label, width=20, anchor="w").pack(side="left")
        entry = tk.Entry(row, textvariable=var)
        entry.pack(side="left", fill="x", expand=True)
        tk.Button(row, text="选择", width=10, command=command).pack(side="left", padx=(8, 0))

        if DND_AVAILABLE and drop_types:
            self._enable_drop(entry, var, drop_types)

    def _add_drop_zone(self, parent, title, var, allowed_exts):
        zone = tk.Label(
            parent,
            text=title,
            relief="groove",
            bd=1,
            height=3,
            anchor="center",
            bg="#F8FAFC",
            fg="#334155",
        )
        zone.pack(fill="x", pady=(4, 8))

        if DND_AVAILABLE:
            try:
                zone.drop_target_register(DND_FILES)
                zone.dnd_bind("<<Drop>>", lambda e: self._handle_drop(e, var, allowed_exts))
                zone.configure(text=title + "\n(可拖拽文件到此区域)")
            except Exception:
                zone.configure(text=title + "\n(拖拽初始化失败，请使用选择按钮)")
        else:
            zone.configure(text=title + "\n(当前环境未启用拖拽，请使用选择按钮)")

    def _enable_drop(self, widget, var_obj, allowed_exts):
        try:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", lambda e: self._handle_drop(e, var_obj, allowed_exts))
        except Exception:
            self._append_log("提示：拖拽功能初始化失败，将继续使用选择按钮。")

    def _handle_drop(self, event, var_obj, allowed_exts):
        raw = event.data.strip()
        path = raw.strip("{}").split(" ")[0].strip("{}")
        ext = os.path.splitext(path)[1].lower()
        if ext not in allowed_exts:
            messagebox.showwarning("格式不支持", f"仅支持: {', '.join(sorted(allowed_exts))}")
            return
        var_obj.set(path)

    def _append_log(self, msg: str):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    @staticmethod
    def _normalize_col_name(name: str) -> str:
        text = str(name)
        text = INVISIBLE_CHAR_PATTERN.sub("", text)
        return text.strip()

    @staticmethod
    def _clean_cell(value):
        text = "" if value is None else str(value)
        text = text.strip().rstrip()
        if len(text) >= 2 and text[0] == '"' and text[-1] == '"':
            text = text[1:-1]
        return text.strip()

    @staticmethod
    def clean_leading_numbering(text):
        s = "" if text is None else str(text)
        s = s.strip()
        if not s:
            return ""
        pattern = r"^\s*(?:\(?\d+\)?|（\d+）)\s*(?:[\.、\):：\-]|-)?\s*"
        s = re.sub(pattern, "", s)
        return s.strip()

    @staticmethod
    def clean_gemini_code_block(content: str) -> str:
        text = content.replace("\r\n", "\n").replace("\r", "\n")
        fence = re.search(r"```(?:tsv)?\s*\n(.*?)\n```", text, re.S | re.I)
        if fence:
            return fence.group(1).strip()
        lines = [ln for ln in text.split("\n") if ln.strip()]
        return "\n".join(lines).strip()

    @staticmethod
    def _default_output_path(input_path: str) -> str:
        base_dir = os.path.dirname(input_path)
        name, _ = os.path.splitext(os.path.basename(input_path))
        return os.path.join(base_dir, f"{name}_extracted.xlsx")

    @staticmethod
    def _default_gemini_output_dir(input_path: str) -> str:
        base_dir = os.path.dirname(input_path)
        return os.path.join(base_dir, "gemini_batches")

    @staticmethod
    def _default_vertical_filename() -> str:
        return f"gemini_result_vertical_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    def _ensure_xlsx_name(self, filename: str) -> str:
        name = filename.strip()
        if not name:
            name = self._default_vertical_filename()
        if not name.lower().endswith(".xlsx"):
            name += ".xlsx"
        return name

    def _create_task(self, task_name, source):
        self.task_counter += 1
        task_id = f"task_{self.task_counter:04d}"
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        task = {
            "id": task_id,
            "task_name": task_name,
            "source": source,
            "status": "处理中",
            "count": "",
            "output": "",
            "created": now,
            "finished": "",
            "remark": "",
        }
        self.tasks.append(task)
        self.task_tree.insert(
            "", "end", iid=task_id,
            values=(task["task_name"], task["source"], task["status"], task["count"], task["output"], task["created"], task["finished"], task["remark"]),
        )
        return task

    def _update_task(self, task, status=None, count=None, output=None, remark=None):
        if status is not None:
            task["status"] = status
        if count is not None:
            task["count"] = count
        if output is not None:
            task["output"] = output
        if remark is not None:
            task["remark"] = remark
        if status in {"已完成", "失败"}:
            task["finished"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        self.task_tree.item(
            task["id"],
            values=(
                task["task_name"],
                task["source"],
                task["status"],
                task["count"],
                task["output"],
                task["created"],
                task["finished"],
                task["remark"],
            ),
        )

    def choose_input_file(self):
        file_path = filedialog.askopenfilename(title="选择输入 Excel 文件", filetypes=[("Excel 文件", "*.xlsx *.xls")])
        if file_path:
            self.input_path_var.set(file_path)
            if not self.output_path_var.get():
                self.output_path_var.set(self._default_output_path(file_path))
            if not self.gemini_output_dir_var.get():
                self.gemini_output_dir_var.set(self._default_gemini_output_dir(file_path))

    def choose_output_file(self):
        file_path = filedialog.asksaveasfilename(
            title="选择输出 Excel 文件", defaultextension=".xlsx", filetypes=[("Excel 文件", "*.xlsx")], initialfile="extracted_result.xlsx"
        )
        if file_path:
            self.output_path_var.set(file_path)

    def choose_gemini_output_dir(self):
        dir_path = filedialog.askdirectory(title="选择 Gemini TXT 输出目录")
        if dir_path:
            self.gemini_output_dir_var.set(dir_path)

    def choose_gemini_source_file(self):
        file_path = filedialog.askopenfilename(
            title="选择 Gemini 结果文件",
            filetypes=[("Gemini 结果", "*.txt *.tsv *.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if file_path:
            self.gemini_source_file_var.set(file_path)

    def choose_vertical_output_dir(self):
        dir_path = filedialog.askdirectory(title="选择纵向 Excel 输出目录")
        if dir_path:
            self.vertical_output_dir_var.set(dir_path)
            if not self.vertical_filename_var.get().strip():
                self.vertical_filename_var.set(self._default_vertical_filename())

    def _find_column(self, columns, target_name: str, case_insensitive: bool = False):
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
        work_df["英语-标题"] = work_df["英语-标题"].fillna("").astype(str)
        work_df["英语-描述"] = work_df["英语-描述"].fillna("").astype(str)
        return work_df

    def run_extract(self):
        task = self._create_task("Excel去重导出", "Excel去重导出")
        input_path = self.input_path_var.get().strip()
        output_path = self.output_path_var.get().strip()

        try:
            if not input_path:
                raise ValueError("请先选择输入 Excel 文件。")
            if not os.path.exists(input_path):
                raise ValueError("输入文件不存在，请重新选择。")
            if not output_path:
                output_path = self._default_output_path(input_path)
                self.output_path_var.set(output_path)

            work_df = self._load_required_columns(input_path)
            dedup_df = work_df.drop_duplicates(subset=["id"], keep="first")

            output_dir = os.path.dirname(output_path)
            if output_dir:
                os.makedirs(output_dir, exist_ok=True)
            dedup_df.to_excel(output_path, index=False)

            self._append_log(f"Excel去重导出完成: {output_path}")
            self._update_task(task, status="已完成", count=str(len(dedup_df)), output=output_path, remark="导出成功")
            messagebox.showinfo("完成", f"Excel 提取完成！\n输出文件：\n{output_path}")
        except Exception as exc:
            self._append_log(f"Excel去重导出失败: {exc}")
            self._update_task(task, status="失败", remark=str(exc))
            messagebox.showerror("错误", str(exc))

    def generate_gemini_batch_txt(self, dedup_df: pd.DataFrame, output_dir: str, batch_size: int):
        os.makedirs(output_dir, exist_ok=True)
        total = len(dedup_df)
        file_count = int(math.ceil(total / batch_size))
        batch_meta = []

        for file_index in range(file_count):
            start = file_index * batch_size
            end = min(start + batch_size, total)
            chunk = dedup_df.iloc[start:end]

            filename = f"gemini_batch_{file_index + 1:03d}.txt"
            file_path = os.path.join(output_dir, filename)
            ids = [str(v).strip() for v in chunk["id"].tolist()]
            ids_line = "IDS: " + ",".join(ids)

            blocks = []
            for _, row in chunk.iterrows():
                block = (
                    "[PRODUCT]\n"
                    f"id: {str(row['id']).strip()}\n"
                    f"原始标题: {str(row['英语-标题']) if pd.notna(row['英语-标题']) else ''}\n"
                    f"原始描述: {str(row['英语-描述']) if pd.notna(row['英语-描述']) else ''}\n"
                    "[/PRODUCT]"
                )
                blocks.append(block)

            content = ids_line + "\n\n" + "\n\n".join(blocks)
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(content)

            batch_meta.append({"file": file_path, "ids": ids})

        return file_count, batch_meta

    def generate_gemini_batches(self):
        task = self._create_task("Gemini批量输入生成", "Gemini批量输入生成")
        input_path = self.input_path_var.get().strip()
        output_dir = self.gemini_output_dir_var.get().strip()
        batch_size_str = self.batch_size_var.get().strip()

        try:
            if not input_path:
                raise ValueError("请先选择输入 Excel 文件。")
            if not os.path.exists(input_path):
                raise ValueError("输入文件不存在，请重新选择。")
            if not output_dir:
                output_dir = self._default_gemini_output_dir(input_path)
                self.gemini_output_dir_var.set(output_dir)

            batch_size = int(batch_size_str or "10")
            if batch_size <= 0:
                raise ValueError("批量大小必须是大于 0 的整数。")

            work_df = self._load_required_columns(input_path)
            dedup_df = work_df.drop_duplicates(subset=["id"], keep="first").reset_index(drop=True)
            if dedup_df.empty:
                raise ValueError("去重后没有可生成的产品数据。")

            file_count, batch_meta = self.generate_gemini_batch_txt(dedup_df, output_dir, batch_size)
            remark = json.dumps(batch_meta[:10], ensure_ascii=False)
            if len(batch_meta) > 10:
                remark += " ..."

            self._append_log(f"Gemini 批量生成完成：共 {file_count} 个 txt")
            self._update_task(task, status="已完成", count=str(len(dedup_df)), output=output_dir, remark=f"共{file_count}个文件; {remark}")
            messagebox.showinfo("完成", f"Gemini 批量输入生成完成！\n共生成 {file_count} 个 txt 文件。")
        except Exception as exc:
            self._append_log(f"Gemini 批量生成失败: {exc}")
            self._update_task(task, status="失败", remark=str(exc))
            messagebox.showerror("错误", str(exc))

    def _parse_tsv_text(self, tsv_text: str) -> pd.DataFrame:
        cleaned = self.clean_gemini_code_block(tsv_text)
        if not cleaned.strip():
            raise ValueError("未检测到可解析的 TSV 内容。")

        # 兼容网页复制后把换行变成字面量 \n 的场景
        cleaned = cleaned.replace("\r\n", "\n").replace("\r", "\n")
        if "\\n" in cleaned:
            cleaned = cleaned.replace("\\n", "\n")

        lines = [ln.rstrip() for ln in cleaned.split("\n") if ln.strip()]
        if not lines:
            raise ValueError("解析不到表头，请确认 TSV 文本不为空。")

        header_line = lines[0].strip()
        headers = [self._normalize_col_name(h) for h in header_line.split("\t")]
        if not headers or headers[0] == "":
            raise ValueError("解析不到表头，请确认第一行为 TSV 表头并使用 Tab 分隔。")

        rows = []
        for line in lines[1:]:
            line_strip = line.strip()
            # 跳过重复出现的表头行
            if line_strip == header_line:
                continue
            if line_strip.startswith("id\t原始标题\t原始描述"):
                continue

            parts = line.rstrip().split("\t")
            if len(parts) < len(headers):
                parts += [""] * (len(headers) - len(parts))
            elif len(parts) > len(headers):
                # 超长时把多余列拼回最后一列，避免整体错位
                parts = parts[: len(headers) - 1] + ["\t".join(parts[len(headers) - 1 :])]

            rows.append(parts)

        if not rows:
            raise ValueError("没有解析到有效数据行，请检查 TSV 内容格式。")

        df = pd.DataFrame(rows, columns=headers)
        df = df.fillna("")
        for col in df.columns:
            df[col] = df[col].apply(self._clean_cell)
        return df

    def read_gemini_result_file(self, input_path: str) -> pd.DataFrame:
        ext = os.path.splitext(input_path)[1].lower()
        if ext in {".txt", ".tsv"}:
            with open(input_path, "r", encoding="utf-8") as f:
                text = f.read()
            return self._parse_tsv_text(text)
        if ext in {".xlsx", ".xls"}:
            df = pd.read_excel(input_path, dtype=str)
            if df.empty:
                raise ValueError("输入文件为空，没有可转换的数据。")
            df.columns = [self._normalize_col_name(c) for c in df.columns]
            df = df.fillna("")
            for col in df.columns:
                df[col] = df[col].apply(self._clean_cell)
            return df
        raise ValueError("仅支持 .txt / .tsv / .xlsx / .xls 文件。")

    def _join_old_spec_fields(self, row: pd.Series, keys: list[str]) -> str:
        vals = []
        for key in keys:
            v = self._clean_cell(row.get(key, ""))
            if v:
                vals.append(v)
        return " || ".join(vals)

    def _format_spec_multiline(self, text: str) -> str:
        raw = self._clean_cell(text)
        if not raw:
            return ""
        parts = [self.clean_leading_numbering(x) for x in re.split(r"\s*\|\|\s*", raw)]
        parts = [x for x in [p.strip() for p in parts] if x]
        return "\n".join([f"{i}. {v}" for i, v in enumerate(parts, start=1)])

    def convert_gemini_result_to_vertical_excel(self, df: pd.DataFrame, output_path: str):
        missing = [f for f in GEMINI_FIELD_ORDER if f not in df.columns]
        if missing:
            self._append_log(f"提示：缺少字段，已按空值处理: {', '.join(missing)}")

        has_new_spec_en = "AI_规格描述_EN" in df.columns
        has_new_spec_cn = "AI_规格描述_CN" in df.columns
        has_old_spec_en = any(c in df.columns for c in OLD_SPEC_EN_FIELDS)
        has_old_spec_cn = any(c in df.columns for c in OLD_SPEC_CN_FIELDS)

        if not has_new_spec_en and has_old_spec_en:
            self._append_log("提示：检测到旧版规格字段 EN_1~EN_5，已自动合并为 AI_规格描述_EN")
        if not has_new_spec_cn and has_old_spec_cn:
            self._append_log("提示：检测到旧版规格字段 CN_1~CN_5，已自动合并为 AI_规格描述_CN")

        records = []
        for _, row in df.iterrows():
            spec_en_source = self._clean_cell(row.get("AI_规格描述_EN", "")) if has_new_spec_en else self._join_old_spec_fields(row, OLD_SPEC_EN_FIELDS)
            spec_cn_source = self._clean_cell(row.get("AI_规格描述_CN", "")) if has_new_spec_cn else self._join_old_spec_fields(row, OLD_SPEC_CN_FIELDS)

            for field in GEMINI_FIELD_ORDER:
                if field == "AI_规格描述_EN":
                    value = self._format_spec_multiline(spec_en_source)
                elif field == "AI_规格描述_CN":
                    value = self._format_spec_multiline(spec_cn_source)
                else:
                    value = row[field] if field in row.index else ""
                    if field in NUMBER_CLEAN_FIELDS:
                        value = self.clean_leading_numbering(value)
                    value = self._clean_cell(value)

                records.append({"字段": field, "内容": value})
            records.append({"字段": "", "内容": ""})

        out_df = pd.DataFrame(records, columns=["字段", "内容"])
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            out_df.to_excel(writer, index=False, sheet_name="Vertical")
            ws = writer.sheets["Vertical"]
            self.apply_vertical_excel_styles(ws)

    def apply_vertical_excel_styles(self, ws):
        field_fill = PatternFill(fill_type="solid", fgColor="DCEBFF")
        id_fill = PatternFill(fill_type="solid", fgColor="C7DFFF")

        ws.freeze_panes = "A2"
        ws.cell(row=1, column=1).font = Font(bold=True)
        ws.cell(row=1, column=2).font = Font(bold=True)
        ws.cell(row=1, column=1).fill = field_fill
        ws.cell(row=1, column=2).fill = field_fill

        for i in range(2, ws.max_row + 1):
            fc = ws.cell(i, 1)
            cc = ws.cell(i, 2)

            fc.font = Font(bold=True)
            fc.fill = field_fill
            fc.alignment = Alignment(vertical="top")
            cc.alignment = Alignment(wrap_text=True, vertical="top")

            if str(fc.value).strip() == "id":
                fc.fill = id_fill
                cc.fill = id_fill
                cc.font = Font(bold=True)

        ws.column_dimensions["A"].width = 28
        ws.column_dimensions["B"].width = 90

    def _build_vertical_output_path(self):
        out_dir = self.vertical_output_dir_var.get().strip()
        if not out_dir:
            raise ValueError("请先选择输出目录。")
        filename = self._ensure_xlsx_name(self.vertical_filename_var.get())
        self.vertical_filename_var.set(filename)
        return os.path.join(out_dir, filename)

    def run_convert_gemini_vertical_from_file(self):
        task = self._create_task("Gemini TSV转纵向表", "文件导入")
        try:
            input_path = self.gemini_source_file_var.get().strip()
            if not input_path:
                raise ValueError("请先选择 Gemini 输入文件。")
            if not os.path.exists(input_path):
                raise ValueError("输入文件不存在。")

            df = self.read_gemini_result_file(input_path)
            output_path = self._build_vertical_output_path()
            self.convert_gemini_result_to_vertical_excel(df, output_path)

            self._append_log(f"转换完成，共处理 {len(df)} 个产品，输出文件：{output_path}")
            self._update_task(task, status="已完成", count=str(len(df)), output=output_path, remark="转换成功")
            messagebox.showinfo("完成", f"转换完成，共处理 {len(df)} 个产品。\n输出文件：\n{output_path}")
        except Exception as exc:
            self._append_log(f"Gemini 纵向转换失败: {exc}")
            self._update_task(task, status="失败", remark=str(exc))
            messagebox.showerror("错误", str(exc))

    def run_convert_gemini_vertical_from_paste(self):
        task = self._create_task("Gemini TSV转纵向表", "粘贴文本")
        try:
            text = self.tsv_paste_text.get("1.0", "end")
            df = self._parse_tsv_text(text)
            output_path = self._build_vertical_output_path()
            self.convert_gemini_result_to_vertical_excel(df, output_path)

            self._append_log(f"粘贴文本转换完成，共处理 {len(df)} 个产品，输出文件：{output_path}")
            self._update_task(task, status="已完成", count=str(len(df)), output=output_path, remark="转换成功")
            messagebox.showinfo("完成", f"转换完成，共处理 {len(df)} 个产品。\n输出文件：\n{output_path}")
        except Exception as exc:
            self._append_log(f"粘贴文本转换失败: {exc}")
            self._update_task(task, status="失败", remark=str(exc))
            messagebox.showerror("错误", str(exc))


def main():
    if DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    ExcelExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
