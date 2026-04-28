# ExcelExtractor（Excel 商品数据提取桌面工具）

桌面工具包含三个功能：

1. Excel 去重提取导出
2. 生成 Gemini 批量输入 TXT（含 IDS 行）
3. Gemini TSV/TXT/XLSX 结果转纵向查看 Excel

## 目录

```text
project/
├─ main.py
├─ ExcelExtractor.spec
├─ build.bat
├─ requirements.txt
└─ README.md
```

## 打包

在 `project` 目录双击 `build.bat`，成功后输出：

```text
dist\ExcelExtractor.exe
```

## 功能1：Excel 去重提取

输入列要求：`id`（大小写不敏感）、`英语-标题`、`英语-描述`。

输出：去重后 Excel（按 `id` 保留首次出现）。

## 功能2：生成 Gemini 批量输入 TXT

- 默认每 10 条产品一个 txt（可在界面修改批量大小）
- 文件名：`gemini_batch_001.txt`、`gemini_batch_002.txt`...
- 每个 txt 顶部新增 `IDS` 行（逗号分隔，不换行）
- `IDS` 行后空一行，再写 `[PRODUCT]` 块

示例：

```text
IDS: 683443087,683443086,683443085

[PRODUCT]
id: 683443087
原始标题: xxx
原始描述: xxx
[/PRODUCT]
```

## 功能3：Gemini 结果转纵向 Excel

### 输入支持

- `.txt` / `.tsv`：按 Tab 读取，首行表头
- `.xlsx` / `.xls`：读取第一个 Sheet
- 自动清洗列名空格与不可见字符
- 若输入带有 ``` 或 ```tsv 代码块包裹，会自动移除包裹标记

### 输出

- 文件名固定：`gemini_result_vertical.xlsx`
- 两列：`字段`、`内容`
- 每个产品为一个纵向区块，产品之间插入空行

字段顺序：

1. 产品ID
2. AI状态（固定“成功”）
3. 原始标题
4. 原始描述
5. 违禁词检验
6. 品牌识别
7. 通用标题 EN
8. 通用标题 CN
9. 关键词 EN
10. 关键词 CN
11. 卖点 EN（1~5 合并到一个单元格，自动编号，跳过空项）
12. 卖点 CN（1~5 合并到一个单元格，自动编号，跳过空项）
13. 规格描述 EN（1~5 合并到一个单元格，自动编号，跳过空项）
14. 规格描述 CN（1~5 合并到一个单元格，自动编号，跳过空项）

### 样式

- 表头加粗
- 字段列加粗 + 浅蓝背景
- 内容列自动换行
- 冻结首行
- “产品ID”行加深底色
- 自动调整列宽与行高

### 缺字段处理

若缺少字段，不崩溃：对应内容留空，并在日志提示缺少字段。
