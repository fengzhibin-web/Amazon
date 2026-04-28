@echo off
setlocal

REM 一键打包：生成单文件、无控制台窗口的 exe
REM 若需要先安装依赖：pip install -r requirements.txt

pyinstaller --noconfirm --clean --onefile --windowed --name ExcelExtractor ^
  --hidden-import=tkinter ^
  --hidden-import=tkinter.filedialog ^
  --hidden-import=tkinter.messagebox ^
  --hidden-import=openpyxl ^
  --hidden-import=pandas ^
  main.py

if %errorlevel% neq 0 (
  echo.
  echo 打包失败，请检查错误日志。
  exit /b %errorlevel%
)

echo.
echo 打包完成！
echo EXE 路径：dist\ExcelExtractor.exe
endlocal
