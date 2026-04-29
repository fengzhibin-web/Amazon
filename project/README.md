# ExcelExtractor（Excel / Gemini 工具）

## 流程1：Excel去重 + Gemini批量输入TXT

输入：原始 Excel（包含 `id / 英语-标题 / 英语-描述`）  
输出：
- 去重 Excel
- `gemini_batch_XXX.txt`（每个文件顶部 `IDS:`，下方 `[PRODUCT]`）

## 流程2：Gemini TSV 转纵向 Excel

支持两种输入方式：
1. 选择 Gemini 结果文件（`.txt/.tsv/.xlsx/.xls`）
2. 直接粘贴 TSV 文本（支持 ```tsv 代码块）

### 新版 Gemini 必需字段

`id`、`品牌识别`、`AI_通用标题_EN`、`AI_通用标题_CN`、
`AI_卖点_EN_1..5`、`AI_卖点_CN_1..5`、`AI_规格描述_EN`、`AI_规格描述_CN`

> 默认不再输出旧版 `原始标题/原始描述/关键词` 字段。

### 可选：输出原始标题和原始描述

勾选 `输出原始标题和原始描述` 后：
- 优先从 batch 源文件 txt 按 id 提取原始标题/原始描述
- 若未提供 batch 源 txt，则尝试从旧 TSV 回退
- 若都没有，留空并在日志提示未匹配 id

### 规格描述显示

`AI_规格描述_EN/CN` 内部使用 ` || ` 分隔多条时，导出到 Excel 会：
- 拆分为多行
- 转成单元格内多行显示
- 每条先执行开头序号清洗（如 1./2./(3)/（4） 等）

### 表头错误检测

若检测到错误表头 `id品牌识别`（缺少 Tab），会直接报错并提示修正提示词。

## 任务列表

包含：任务名、输入来源、状态、产品数量、输出文件、创建时间、完成时间、备注。

## 拖拽支持

若安装 `tkinterdnd2`：
- 流程1支持拖拽 `.xlsx/.xls`
- 流程2支持拖拽 Gemini 结果文件 `.txt/.tsv/.xlsx/.xls`
- Batch 源文件支持拖拽 `.txt`

## 打包

```bat
build.bat
```

核心命令包含：

```bat
python -m PyInstaller --onefile --windowed --name ExcelExtractor --collect-all tkinterdnd2 main.py
```
