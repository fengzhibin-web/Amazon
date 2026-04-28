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

## 为什么双击 `build.bat` 会“没反应”？

最常见原因：

1. 电脑没有安装 Python（或未加入 PATH）
2. pip 安装依赖失败后窗口瞬间关闭
3. 杀毒软件拦截了 PyInstaller

本项目的新版 `build.bat` 已做了改进：

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
   - 某些环境下可避免运行时缺组件

---

## 一键打包 Windows EXE

在 `project` 目录下双击 `build.bat`，或在命令行执行：

```bat
build.bat
```

脚本内部核心命令为：

```bat
pyinstaller --onefile --windowed --name ExcelExtractor main.py
```

并额外补充了常见隐藏依赖 `hidden-import`，降低缺模块风险。

打包成功后文件位置：

```text
dist\ExcelExtractor.exe
```

---

## 给最终用户（无需 Python）

将 `dist\ExcelExtractor.exe` 发给最终用户即可。用户电脑 **不需要安装 Python**：

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
