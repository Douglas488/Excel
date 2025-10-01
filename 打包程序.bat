@echo off
chcp 65001 >nul
setlocal ENABLEDELAYEDEXPANSION

REM 进入脚本所在目录
cd /d %~dp0

REM 选择图标：优先 app.ico，回退 app_icon.ico.ico
set ICON=app.ico
if not exist "app.ico" (
    if exist "app_icon.ico.ico" (
        set ICON=app_icon.ico.ico
    ) else (
        echo 未找到 app.ico 或 app_icon.ico.ico，继续使用无图标构建。
        set ICON=
    )
)

REM 检查 PyInstaller 是否安装
where pyinstaller >nul 2>nul
if errorlevel 1 (
    echo 未检测到 PyInstaller，请先安装：
    echo   pip install pyinstaller
    echo （如需图片缩放功能，请一并安装：pip install pillow）
    pause
    goto :eof
)

REM 清理上次构建
if exist build rd /s /q build
if exist dist rd /s /q dist
if exist excel_processor.spec del /f /q excel_processor.spec

REM 组合 --icon 参数
set ICON_ARG=
if not "%ICON%"=="" set ICON_ARG=-i "%ICON%"

REM 打包静态资源（根目录与 img/ 目录都尝试）
set DATA_ARGS=--add-data "01.png;." --add-data "02.png;." --add-data "03.png;." --add-data "使用说明.md;."
if exist "img\01.png" set DATA_ARGS=%DATA_ARGS% --add-data "img\01.png;img"
if exist "img\02.png" set DATA_ARGS=%DATA_ARGS% --add-data "img\02.png;img"
if exist "img\03.png" set DATA_ARGS=%DATA_ARGS% --add-data "img\03.png;img"

REM 打包：单文件、无控制台、中文名称
pyinstaller --noconfirm --clean -F -w %ICON_ARG% --name "Excel数据处理工具" ^
  %DATA_ARGS% ^
  --collect-all openpyxl --collect-all pandas ^
  excel_processor.py

if errorlevel 1 (
    echo 构建失败，请检查报错信息。
    pause
    goto :eof
)

REM 输出位置提示
if exist "dist\Excel数据处理工具.exe" (
    echo 构建成功：dist\Excel数据处理工具.exe
) else (
    echo 构建完成，请在 dist 目录查看生成的 EXE。
)

pause
endlocal
