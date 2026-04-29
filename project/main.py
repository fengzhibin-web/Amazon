import json
import math
import os
import re
import subprocess
import sys
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill

APP_TITLE = "Excel 商品数据提取工具"
WINDOW_SIZE = "1180x760"
INVISIBLE_CHAR_PATTERN = re.compile(r"[\u200b\u200c\u200d\ufeff]")

GEMINI_FIELD_ORDER = [
    "id", "原始标题", "原始描述", "违禁词检验", "品牌识别",
    "AI_通用标题_EN", "AI_通用标题_CN", "AI_关键词_EN", "AI_关键词_CN",
    "AI_卖点_EN_1", "AI_卖点_EN_2", "AI_卖点_EN_3", "AI_卖点_EN_4", "AI_卖点_EN_5",
    "AI_卖点_CN_1", "AI_卖点_CN_2", "AI_卖点_CN_3", "AI_卖点_CN_4", "AI_卖点_CN_5",
    "AI_规格描述_EN", "AI_规格描述_CN",
]

OLD_SPEC_EN_FIELDS = [f"AI_规格描述_EN_{i}" for i in range(1, 6)]
OLD_SPEC_CN_FIELDS = [f"AI_规格描述_CN_{i}" for i in range(1, 6)]

NUMBER_CLEAN_FIELDS = {
    "AI_卖点_EN_1", "AI_卖点_EN_2", "AI_卖点_EN_3", "AI_卖点_EN_4", "AI_卖点_EN_5",
    "AI_卖点_CN_1", "AI_卖点_CN_2", "AI_卖点_CN_3", "AI_卖点_CN_4", "AI_卖点_CN_5",
    "AI_规格描述_EN", "AI_规格描述_CN",
}


class ExcelExtractorApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry(WINDOW_SIZE)

        self.input_path_var = tk.StringVar()
        self.batch_size_var = tk.StringVar(value="10")

        self.dedup_output_dir_var = tk.StringVar()
        self.dedup_filename_var = tk.StringVar(value="extracted_result.xlsx")

        self.batch_output_dir_var = tk.StringVar()

        self.vertical_output_dir_var = tk.StringVar()
        self.vertical_filename_var = tk.StringVar(value=self._default_vertical_filename())

        self.tasks = []
        self.task_counter = 0
        self._log_link_counter = 0

        self._build_ui()

    def _build_ui(self):
        top_frame = tk.Frame(self.root)
        top_frame.pack(fill="both", expand=True, padx=12, pady=8)

        tk.Label(
            top_frame,
            text="流程1：原始Excel -> 去重导出 + Gemini批量TXT | 流程2：粘贴Gemini TSV -> 纵向Excel",
            font=("Microsoft YaHei", 10, "bold"),
            anchor="w",
        ).pack(fill="x", pady=(0, 6))

        panes = tk.PanedWindow(top_frame, orient="horizontal", sashrelief="raised")
        panes.pack(fill="both", expand=True)

        left = tk.Frame(panes)
        right = tk.Frame(panes)
        panes.add(left, minsize=740)
        panes.add(right, minsize=420)

        self._build_left_area(left)
        self._build_task_area(right)

    def _build_left_area(self, parent):
        notebook = ttk.Notebook(parent)
        notebook.pack(fill="both", expand=True)

        tab_step1 = tk.Frame(notebook)
        tab_step2 = tk.Frame(notebook)
        notebook.add(tab_step1, text="流程1：去重+切分")
        notebook.add(tab_step2, text="流程2：Gemini转纵向")

        self._build_tab_step1(tab_step1)
        self._build_tab_step2(tab_step2)

        tk.Label(parent, text="运行信息（可点击“打开目录”快速跳转）", anchor="w").pack(fill="x", pady=(8, 4))
        self.log_text = tk.Text(parent, height=12, wrap="word", state="disabled")
        self.log_text.pack(fill="both", expand=True)

    def _build_tab_step1(self, parent):
        tutorial = (
            "流程1教程：\n"
            "1) 选择原始Excel。\n"
            "2) 配置批量大小（一个txt包含多少个ID）。\n"
            "3) 配置3个输出目录：去重结果 / 批量TXT / 纵向Excel（供流程2默认使用）。\n"
            "4) 点击“运行流程1”。"
        )
        tk.Label(parent, text=tutorial, justify="left", fg="#334155").pack(fill="x", pady=(0, 8))

        self._row_file(parent, "原始 Excel:", self.input_path_var, self.choose_input_file)

        row = tk.Frame(parent)
        row.pack(fill="x", pady=4)
        tk.Label(row, text="批量大小:", width=20, anchor="w").pack(side="left")
        tk.Entry(row, textvariable=self.batch_size_var, width=10).pack(side="left")
        tk.Label(row, text="（一个文件包含多少ID）", fg="#475569").pack(side="left", padx=(8, 0))

        self._row_dir(parent, "去重输出目录:", self.dedup_output_dir_var, self.choose_dedup_output_dir)
        self._row_file_name(parent, "去重输出文件名:", self.dedup_filename_var)

        self._row_dir(parent, "批量TXT输出目录:", self.batch_output_dir_var, self.choose_batch_output_dir)

        self._row_dir(parent, "纵向Excel输出目录:", self.vertical_output_dir_var, self.choose_vertical_output_dir)
        self._row_file_name(parent, "纵向输出文件名:", self.vertical_filename_var)

        tk.Button(parent, text="运行流程1（去重+切分）", width=24, bg="#0ea5e9", fg="white", command=self.run_step1_pipeline).pack(
            anchor="w", pady=10
        )

    def _build_tab_step2(self, parent):
        tutorial = (
            "流程2教程：\n"
            "1) 把网页端 Gemini 返回的 TSV 文本粘贴到下方。\n"
            "2) 确认纵向输出目录和文件名。\n"
            "3) 点击“运行流程2”。"
        )
        tk.Label(parent, text=tutorial, justify="left", fg="#334155").pack(fill="x", pady=(0, 8))

        self._row_dir(parent, "纵向Excel输出目录:", self.vertical_output_dir_var, self.choose_vertical_output_dir)
        self._row_file_name(parent, "纵向输出文件名:", self.vertical_filename_var)

        tk.Label(parent, text="粘贴 Gemini TSV 文本（支持 ```tsv 包裹）:", anchor="w").pack(fill="x", pady=(8, 4))
        self.tsv_paste_text = tk.Text(parent, height=14, wrap="word")
        self.tsv_paste_text.pack(fill="both", expand=True)

        tk.Button(parent, text="运行流程2（粘贴文本转纵向）", width=26, bg="#9333ea", fg="white", command=self.run_convert_gemini_vertical_from_paste).pack(
            anchor="w", pady=10
        )

    def _build_task_area(self, parent):
        tk.Label(parent, text="任务列表", font=("Microsoft YaHei", 10, "bold")).pack(anchor="w", pady=(0, 6))
        columns = ("task_name", "source", "status", "count", "output", "created", "finished", "remark")
        tree = ttk.Treeview(parent, columns=columns, show="headings", height=30)
        headers = {
            "task_name": "任务名", "source": "输入来源", "status": "状态", "count": "产品数量",
            "output": "输出文件", "created": "创建时间", "finished": "完成时间", "remark": "备注",
        }
        widths = {"task_name": 120, "source": 100, "status": 70, "count": 70, "output": 150, "created": 130, "finished": 130, "remark": 250}
        for col in columns:
            tree.heading(col, text=headers[col])
            tree.column(col, width=widths[col], anchor="w")
        tree.pack(fill="both", expand=True)
        self.task_tree = tree

    def _row_file(self, parent, label, var, cmd):
        row = tk.Frame(parent)
        row.pack(fill="x", pady=4)
        tk.Label(row, text=label, width=20, anchor="w").pack(side="left")
        tk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True)
        tk.Button(row, text="选择", width=10, command=cmd).pack(side="left", padx=(8, 0))

    def _row_dir(self, parent, label, var, cmd):
        self._row_file(parent, label, var, cmd)

    def _row_file_name(self, parent, label, var):
        row = tk.Frame(parent)
        row.pack(fill="x", pady=4)
        tk.Label(row, text=label, width=20, anchor="w").pack(side="left")
        tk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True)

    def _append_log(self, msg: str, path: str | None = None):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")
        if path:
            self._log_link_counter += 1
            tag = f"link_{self._log_link_counter}"
            self.log_text.insert("end", "  [打开目录]", tag)
            self.log_text.tag_config(tag, foreground="#2563eb", underline=True)
            self.log_text.tag_bind(tag, "<Button-1>", lambda e, p=path: self._open_folder(p))
        self.log_text.insert("end", "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _open_folder(self, path: str):
        target = path if os.path.isdir(path) else os.path.dirname(path)
        if not target:
            return
        try:
            if os.name == "nt":
                os.startfile(target)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", target])
            else:
                subprocess.Popen(["xdg-open", target])
        except Exception as exc:
            messagebox.showerror("错误", f"无法打开目录: {exc}")

    @staticmethod
    def _normalize_col_name(name: str) -> str:
        return INVISIBLE_CHAR_PATTERN.sub("", str(name)).strip()

    @staticmethod
    def _clean_cell(value):
        text = "" if value is None else str(value)
        text = text.strip().rstrip()
        if len(text) >= 2 and text[0] == '"' and text[-1] == '"':
            text = text[1:-1]
        return text.strip()

    @staticmethod
    def clean_leading_numbering(text):
        s = "" if text is None else str(text).strip()
        if not s:
            return ""
        return re.sub(r"^\s*(?:\(?\d+\)?|（\d+）)\s*(?:[\.、\):：\-]|-)?\s*", "", s).strip()

    @staticmethod
    def clean_gemini_code_block(content: str) -> str:
        text = content.replace("\r\n", "\n").replace("\r", "\n")
        fence = re.search(r"```(?:tsv)?\s*\n(.*?)\n```", text, re.S | re.I)
        if fence:
            return fence.group(1).strip()
        return "\n".join([ln for ln in text.split("\n") if ln.strip()]).strip()

    @staticmethod
    def _default_vertical_filename() -> str:
        return f"gemini_result_vertical_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    def _ensure_xlsx_name(self, filename: str) -> str:
        name = filename.strip() or self._default_vertical_filename()
        return name if name.lower().endswith(".xlsx") else name + ".xlsx"

    def _create_task(self, task_name, source):
        self.task_counter += 1
        task = {
            "id": f"task_{self.task_counter:04d}",
            "task_name": task_name,
            "source": source,
            "status": "处理中",
            "count": "",
            "output": "",
            "created": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "finished": "",
            "remark": "",
        }
        self.tasks.append(task)
        self.task_tree.insert("", "end", iid=task["id"], values=(task["task_name"], task["source"], task["status"], task["count"], task["output"], task["created"], task["finished"], task["remark"]))
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
        self.task_tree.item(task["id"], values=(task["task_name"], task["source"], task["status"], task["count"], task["output"], task["created"], task["finished"], task["remark"]))

    def choose_input_file(self):
        file_path = filedialog.askopenfilename(title="选择原始 Excel", filetypes=[("Excel 文件", "*.xlsx *.xls")])
        if file_path:
            self.input_path_var.set(file_path)
            base_dir = os.path.dirname(file_path)
            if not self.dedup_output_dir_var.get().strip():
                self.dedup_output_dir_var.set(base_dir)
            if not self.batch_output_dir_var.get().strip():
                self.batch_output_dir_var.set(os.path.join(base_dir, "gemini_batches"))
            if not self.vertical_output_dir_var.get().strip():
                self.vertical_output_dir_var.set(base_dir)

    def choose_dedup_output_dir(self):
        dir_path = filedialog.askdirectory(title="选择去重输出目录")
        if dir_path:
            self.dedup_output_dir_var.set(dir_path)

    def choose_batch_output_dir(self):
        dir_path = filedialog.askdirectory(title="选择批量TXT输出目录")
        if dir_path:
            self.batch_output_dir_var.set(dir_path)

    def choose_vertical_output_dir(self):
        dir_path = filedialog.askdirectory(title="选择纵向Excel输出目录")
        if dir_path:
            self.vertical_output_dir_var.set(dir_path)
            if not self.vertical_filename_var.get().strip():
                self.vertical_filename_var.set(self._default_vertical_filename())

    def _find_column(self, columns, target_name: str, case_insensitive: bool = False):
        if case_insensitive:
            lookup = {str(c).strip().lower(): c for c in columns}
            return lookup.get(target_name.strip().lower())
        return {str(c).strip(): c for c in columns}.get(target_name.strip())

    def _load_required_columns(self, input_path: str) -> pd.DataFrame:
        df = pd.read_excel(input_path)
        if df.empty:
            raise ValueError("输入 Excel 为空，没有可处理的数据。")
        id_col = self._find_column(df.columns, "id", case_insensitive=True)
        title_col = self._find_column(df.columns, "英语-标题")
        desc_col = self._find_column(df.columns, "英语-描述")
        missing = [n for n, c in [("id", id_col), ("英语-标题", title_col), ("英语-描述", desc_col)] if c is None]
        if missing:
            raise ValueError(f"缺少必要列: {', '.join(missing)}")
        work_df = df[[id_col, title_col, desc_col]].copy()
        work_df.columns = ["id", "英语-标题", "英语-描述"]
        work_df["id"] = work_df["id"].astype(str).str.strip()
        work_df = work_df[work_df["id"] != ""].copy()
        work_df["英语-标题"] = work_df["英语-标题"].fillna("").astype(str)
        work_df["英语-描述"] = work_df["英语-描述"].fillna("").astype(str)
        return work_df

    def generate_gemini_batch_txt(self, dedup_df: pd.DataFrame, output_dir: str, batch_size: int):
        os.makedirs(output_dir, exist_ok=True)
        file_count = int(math.ceil(len(dedup_df) / batch_size))
        meta = []
        for i in range(file_count):
            chunk = dedup_df.iloc[i * batch_size: min((i + 1) * batch_size, len(dedup_df))]
            path = os.path.join(output_dir, f"gemini_batch_{i + 1:03d}.txt")
            ids = [str(v).strip() for v in chunk["id"].tolist()]
            blocks = []
            for _, row in chunk.iterrows():
                blocks.append("[PRODUCT]\n"
                              f"id: {str(row['id']).strip()}\n"
                              f"原始标题: {str(row['英语-标题']) if pd.notna(row['英语-标题']) else ''}\n"
                              f"原始描述: {str(row['英语-描述']) if pd.notna(row['英语-描述']) else ''}\n"
                              "[/PRODUCT]")
            content = "IDS: " + ",".join(ids) + "\n\n" + "\n\n".join(blocks)
            with open(path, "w", encoding="utf-8") as f:
                f.write(content)
            meta.append({"file": path, "ids": ids})
        return file_count, meta

    def run_step1_pipeline(self):
        task = self._create_task("流程1：去重+切分", "原始Excel")
        try:
            input_path = self.input_path_var.get().strip()
            if not input_path or not os.path.exists(input_path):
                raise ValueError("请先选择有效的原始 Excel 文件。")

            batch_size = int(self.batch_size_var.get().strip() or "10")
            if batch_size <= 0:
                raise ValueError("批量大小必须是大于 0 的整数。")

            dedup_dir = self.dedup_output_dir_var.get().strip()
            batch_dir = self.batch_output_dir_var.get().strip()
            vertical_dir = self.vertical_output_dir_var.get().strip()
            if not dedup_dir or not batch_dir or not vertical_dir:
                raise ValueError("请配置三个输出目录：去重输出、批量TXT输出、纵向Excel输出。")

            dedup_name = self._ensure_xlsx_name(self.dedup_filename_var.get())
            self.dedup_filename_var.set(dedup_name)
            dedup_path = os.path.join(dedup_dir, dedup_name)

            work_df = self._load_required_columns(input_path)
            dedup_df = work_df.drop_duplicates(subset=["id"], keep="first").reset_index(drop=True)
            if dedup_df.empty:
                raise ValueError("去重后没有可输出的数据。")

            os.makedirs(dedup_dir, exist_ok=True)
            dedup_df.to_excel(dedup_path, index=False)
            self._append_log(f"去重导出完成: {dedup_path}", dedup_path)

            file_count, batch_meta = self.generate_gemini_batch_txt(dedup_df, batch_dir, batch_size)
            self._append_log(f"批量TXT完成: {batch_dir}（共 {file_count} 个文件）", batch_dir)

            remark = json.dumps(batch_meta[:10], ensure_ascii=False)
            if len(batch_meta) > 10:
                remark += " ..."
            self._update_task(task, status="已完成", count=str(len(dedup_df)), output=dedup_path, remark=f"TXT共{file_count}个; {remark}")
            messagebox.showinfo("完成", f"流程1完成：\n去重文件：{dedup_path}\n批量目录：{batch_dir}")
        except Exception as exc:
            self._append_log(f"流程1失败: {exc}")
            self._update_task(task, status="失败", remark=str(exc))
            messagebox.showerror("错误", str(exc))

    def _parse_tsv_text(self, tsv_text: str) -> pd.DataFrame:
        cleaned = self.clean_gemini_code_block(tsv_text)
        if not cleaned.strip():
            raise ValueError("未检测到可解析的 TSV 内容。")
        cleaned = cleaned.replace("\r\n", "\n").replace("\r", "\n")
        if "\\n" in cleaned:
            cleaned = cleaned.replace("\\n", "\n")

        lines = [ln.rstrip() for ln in cleaned.split("\n") if ln.strip()]
        if not lines:
            raise ValueError("解析不到表头，请确认 TSV 文本不为空。")

        header = [self._normalize_col_name(h) for h in lines[0].strip().split("\t")]
        if not header or header[0] == "":
            raise ValueError("解析不到表头，请确认第一行为 TSV 表头并使用 Tab 分隔。")

        rows = []
        for line in lines[1:]:
            if line.strip() == lines[0].strip() or line.strip().startswith("id\t原始标题\t原始描述"):
                continue
            parts = line.rstrip().split("\t")
            if len(parts) < len(header):
                parts += [""] * (len(header) - len(parts))
            elif len(parts) > len(header):
                parts = parts[:len(header) - 1] + ["\t".join(parts[len(header) - 1:])]
            rows.append(parts)
        if not rows:
            raise ValueError("没有解析到有效数据行，请检查 TSV 内容格式。")

        df = pd.DataFrame(rows, columns=header).fillna("")
        for col in df.columns:
            df[col] = df[col].apply(self._clean_cell)
        return df

    def _join_old_spec_fields(self, row: pd.Series, keys: list[str]) -> str:
        vals = [self._clean_cell(row.get(k, "")) for k in keys]
        vals = [v for v in vals if v]
        return " || ".join(vals)

    def _format_spec_multiline(self, text: str) -> str:
        raw = self._clean_cell(text)
        if not raw:
            return ""
        parts = [self.clean_leading_numbering(x) for x in re.split(r"\s*\|\|\s*", raw)]
        parts = [p.strip() for p in parts if p.strip()]
        return "\n".join([f"{i}. {v}" for i, v in enumerate(parts, start=1)])

    def convert_gemini_result_to_vertical_excel(self, df: pd.DataFrame, output_path: str):
        missing = [f for f in GEMINI_FIELD_ORDER if f not in df.columns]
        if missing:
            self._append_log(f"提示：缺少字段，已按空值处理: {', '.join(missing)}")

        has_new_en = "AI_规格描述_EN" in df.columns
        has_new_cn = "AI_规格描述_CN" in df.columns
        has_old_en = any(c in df.columns for c in OLD_SPEC_EN_FIELDS)
        has_old_cn = any(c in df.columns for c in OLD_SPEC_CN_FIELDS)

        records = []
        for _, row in df.iterrows():
            spec_en = self._clean_cell(row.get("AI_规格描述_EN", "")) if has_new_en else self._join_old_spec_fields(row, OLD_SPEC_EN_FIELDS)
            spec_cn = self._clean_cell(row.get("AI_规格描述_CN", "")) if has_new_cn else self._join_old_spec_fields(row, OLD_SPEC_CN_FIELDS)
            for field in GEMINI_FIELD_ORDER:
                if field == "AI_规格描述_EN":
                    value = self._format_spec_multiline(spec_en)
                elif field == "AI_规格描述_CN":
                    value = self._format_spec_multiline(spec_cn)
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
        ws.cell(1, 1).font = Font(bold=True)
        ws.cell(1, 2).font = Font(bold=True)
        ws.cell(1, 1).fill = field_fill
        ws.cell(1, 2).fill = field_fill

        for i in range(2, ws.max_row + 1):
            fc = ws.cell(i, 1)
            cc = ws.cell(i, 2)
            fc.font = Font(bold=True)
            fc.fill = field_fill
            fc.alignment = Alignment(vertical="top")
            cc.alignment = Alignment(wrap_text=True, vertical="top")
            text = str(cc.value or "")
            lines = text.count("\n") + 1
            ws.row_dimensions[i].height = max(20, min(140, lines * 18))
            if str(fc.value).strip() == "id":
                fc.fill = id_fill
                cc.fill = id_fill
                cc.font = Font(bold=True)

        ws.column_dimensions["A"].width = 28
        ws.column_dimensions["B"].width = 90

    def _build_vertical_output_path(self):
        out_dir = self.vertical_output_dir_var.get().strip()
        if not out_dir:
            raise ValueError("请先选择纵向Excel输出目录。")
        filename = self._ensure_xlsx_name(self.vertical_filename_var.get())
        self.vertical_filename_var.set(filename)
        return os.path.join(out_dir, filename)

    def run_convert_gemini_vertical_from_paste(self):
        task = self._create_task("流程2：Gemini转纵向", "粘贴文本")
        try:
            text = self.tsv_paste_text.get("1.0", "end")
            df = self._parse_tsv_text(text)
            out_path = self._build_vertical_output_path()
            self.convert_gemini_result_to_vertical_excel(df, out_path)
            self._append_log(f"流程2完成，共处理 {len(df)} 个产品: {out_path}", out_path)
            self._update_task(task, status="已完成", count=str(len(df)), output=out_path, remark="转换成功")
            messagebox.showinfo("完成", f"流程2完成，共处理 {len(df)} 个产品。\n输出文件：\n{out_path}")
        except Exception as exc:
            self._append_log(f"流程2失败: {exc}")
            self._update_task(task, status="失败", remark=str(exc))
            messagebox.showerror("错误", str(exc))


def main():
    root = tk.Tk()
    ExcelExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
