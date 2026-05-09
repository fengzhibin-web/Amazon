# ExcelExtractor（Excel / Gemini 工具）

## 流程1：Excel去重 + Gemini批量输入TXT

输入：原始 Excel（包含 `id / 英语-标题 / 英语-描述`）  
输出：
- 去重 Excel
- `gemini_batch_XXX.txt`（每个文件顶部 `IDS:`，下方 `[PRODUCT]`）

## 流程2：粘贴 Gemini TSV 转纵向 Excel

当前页面仅保留一个 TSV 粘贴框（不再提供 Gemini 结果文件选择和拖拽框）。

### 可选输出开关（分离）

- `输出原始标题`
- `输出原始描述`

可分别勾选。勾选后会尝试从 `Batch源文件TXT` 按 id 回填对应内容。

### 新版 Gemini 必需字段

`id`、`品牌识别`、`AI_通用标题_EN`、`AI_通用标题_CN`、
`AI_卖点_EN_1..5`、`AI_卖点_CN_1..5`、`AI_规格描述_EN`、`AI_规格描述_CN`

### 规格描述显示

`AI_规格描述_EN/CN` 内部使用 ` || ` 分隔多条时，导出到 Excel 会：
- 拆分为多行
- 每条先执行开头序号清洗（如 1./2./(3)/（4） 等）
- 不再强制重新编号，避免重复序号

### 行高策略

纵向表行高改为更紧凑估算（约 13.5 * 可视行数），并保留自动换行，减少过高行的问题。

### 表头错误检测

若检测到错误表头 `id品牌识别`（缺少 Tab），会直接报错并提示修正提示词。

## 任务列表

包含：任务名、输入来源、状态、产品数量、输出文件、创建时间、完成时间、备注。

## 拖拽支持

若安装 `tkinterdnd2`：
- 流程1支持拖拽 `.xlsx/.xls`
- Batch 源文件支持拖拽 `.txt`

## 打包

```bat
build.bat
```

核心命令包含：

```bat
python -m PyInstaller --onefile --windowed --name ExcelExtractor --collect-all tkinterdnd2 main.py
```

### 导出前文本清洗

所有写入 Excel 的字段在导出前会统一执行清洗：
- 清理 `[cite: x]`、`[cite start]`、`[cite end]`、`[citation needed]`、行尾 `[1]` 等引用残留
- 清理包裹性首尾双引号
- 统一换行并去除首尾/中间空白行（保留有效多行）
- 不压缩为单行

日志会输出清洗统计：引用标记、尾部空行、包裹性双引号。
