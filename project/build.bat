@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul
cd /d "%~dp0"

echo =========================================
echo ExcelExtractor 一键打包脚本
echo 当前目录: %CD%
echo =========================================

echo.
echo [1/4] 检查 Python 启动器...
set "PY_CMD="
where py >nul 2>nul
if %errorlevel%==0 (
    set "PY_CMD=py -3"
) else (
    where python >nul 2>nul
    if %errorlevel%==0 (
        set "PY_CMD=python"
    )
)

if "%PY_CMD%"=="" (
    echo [错误] 未检测到 Python。
    echo 请先安装 Python 3.10+ 并勾选 "Add Python to PATH"，然后重试。
    echo 下载: https://www.python.org/downloads/windows/
    echo.
    pause
    exit /b 1
)

echo [OK] 使用命令: %PY_CMD%

echo.
echo [2/4] 安装/更新打包依赖...
%PY_CMD% -m pip install --upgrade pip
if errorlevel 1 goto :pip_fail

%PY_CMD% -m pip install -r requirements.txt
if errorlevel 1 goto :pip_fail

echo.
echo [3/4] 开始打包（单文件 + 无控制台）...
%PY_CMD% -m PyInstaller --noconfirm --clean --onefile --windowed --name ExcelExtractor ^
  --collect-all tkinterdnd2 ^
  --hidden-import=tkinter ^
  --hidden-import=tkinter.filedialog ^
  --hidden-import=tkinter.messagebox ^
  --hidden-import=openpyxl ^
  --hidden-import=pandas ^
  --hidden-import=tkinterdnd2 ^
  main.py
if errorlevel 1 goto :build_fail

echo.
echo [4/4] 打包完成。
echo EXE 路径: %CD%\dist\ExcelExtractor.exe
if exist "%CD%\dist\ExcelExtractor.exe" (
    echo [OK] 文件已生成。
) else (
    echo [警告] 命令成功但未找到 EXE，请检查输出日志。
)

echo.
echo 按任意键退出...
pause >nul
exit /b 0

:pip_fail
echo.
echo [错误] 依赖安装失败。请检查网络、权限，或手动执行:
echo   %PY_CMD% -m pip install -r requirements.txt
echo.
pause
exit /b 2

:build_fail
echo.
echo [错误] 打包失败。常见原因：
echo - 缺少 Visual C++ 运行库
echo - 杀毒软件拦截
echo - 路径含特殊字符/权限不足

echo 可手动重试命令：
echo   %PY_CMD% -m PyInstaller --onefile --windowed --name ExcelExtractor main.py

echo.
pause
exit /b 3
