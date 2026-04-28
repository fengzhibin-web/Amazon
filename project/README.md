# ExcelExtractor（Excel 商品数据提取桌面工具）

一个可打包为 **Windows 单文件 EXE** 的 Tkinter 桌面工具，包含两个功能：

1. **Excel 去重提取**：按 `id` 去重，提取 `英语-标题`、`英语-描述`，导出新 Excel。
2. **Gemini 批量输入生成**：将去重后的产品按批次写入多个 txt 文件。

---

## 目录结构

```text
project/
├─ main.py
├─ ExcelExtractor.spec
├─ build.bat
├─ requirements.txt
└─ README.md
```

---

## 为什么双击 `build.bat` 会“没反应”？

最常见原因：

1. 电脑没有安装 Python（或未加入 PATH）
2. pip 安装依赖失败后窗口瞬间关闭
3. 杀毒软件拦截了 PyInstaller

本项目的 `build.bat` 已做改进：

- 自动检测 `py -3` / `python`
- 自动安装依赖
- 出错时会停在窗口并显示错误提示，不会一闪而过

---

## 打包前你需要安装什么？

> 只在“打包机”上需要安装一次。最终使用 exe 的用户不需要安装 Python。

1. 安装 Python 3.10+（推荐 3.10~3.12）
   - 下载：https://www.python.org/downloads/windows/
   - 安装时勾选 **Add Python to PATH**
2. （可选）安装 Microsoft Visual C++ Redistributable

---

## 一键打包 Windows EXE

在 `project` 目录下双击 `build.bat`，或在命令行执行：

```bat
build.bat
```

打包成功后文件位置：

```text
dist\ExcelExtractor.exe
```

---

## 功能一：Excel 去重提取

输入 Excel 需包含列：

- `id`（大小写不敏感）
- `英语-标题`
- `英语-描述`

处理规则：

- 按 `id` 去重（保留首次出现）
- 导出列：`id`、`英语-标题`、`英语-描述`

---

## 功能二：生成 Gemini 批量输入文本

### 输入

- 去重后的 Excel（也可直接用原始 Excel，程序会先按 `id` 去重）
- 列必须包含：`id`、`英语-标题`、`英语-描述`

### 界面项

- `批量大小` 输入框（默认 `10`）
- `Gemini 输出目录` 选择框
- `生成 Gemini 批量输入` 按钮

### 输出规则

- 每 `N` 条产品（N=批量大小）生成一个 txt
- 文件名：
  - `gemini_batch_001.txt`
  - `gemini_batch_002.txt`
  - `gemini_batch_003.txt`
- 每条产品格式：

```text
[PRODUCT]
id: xxx
原始标题: xxx
原始描述: xxx
[/PRODUCT]
```

若 `英语-描述` 为空，则 `原始描述:` 留空。

---

## 给最终用户（无需 Python）

将 `dist\ExcelExtractor.exe` 发给最终用户即可。用户电脑 **不需要安装 Python**：

1. 双击 `ExcelExtractor.exe`
2. 选择输入 Excel
3. 选择输出路径（Excel 功能）或输出目录（Gemini 功能）
4. 点击对应按钮运行
