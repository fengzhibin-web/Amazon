# ExcelExtractor（Excel 商品数据提取桌面工具）

一个可打包为 **Windows 单文件 EXE** 的 Tkinter 桌面工具，用于：

- 选择 Excel 文件
- 按 `id` 去重
- 提取 `英语-标题`、`英语-描述`
- 导出新 Excel

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

## 本地开发运行

> 仅用于开发调试；最终用户可直接使用打包后的 exe，无需安装 Python。

```bash
pip install -r requirements.txt
python main.py
```

---

## 一键打包 Windows EXE

在 `project` 目录下双击 `build.bat`，或在命令行执行：

```bat
build.bat
```

`build.bat` 内已包含核心命令（单文件 + 无控制台窗口）：

```bat
pyinstaller --onefile --windowed --name ExcelExtractor main.py
```

并额外补充了常见隐藏依赖 `hidden-import`，减少打包后运行缺模块的风险。

打包完成后文件位置：

```text
dist\ExcelExtractor.exe
```

---

## 给最终用户

将 `ExcelExtractor.exe` 发送给目标用户即可。用户电脑 **不需要安装 Python**：

1. 双击 `ExcelExtractor.exe`
2. 选择输入 Excel
3. 选择输出路径
4. 点击“开始提取”

---

## Excel 列要求

输入文件中必须包含以下列：

- `id`（大小写不敏感，例如 `ID` 也可识别）
- `英语-标题`
- `英语-描述`

输出文件列为：

- `id`
- `英语-标题`
- `英语-描述`

去重策略：按 `id` 去重，保留首次出现的记录。
