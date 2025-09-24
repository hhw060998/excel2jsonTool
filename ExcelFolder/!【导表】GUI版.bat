@echo off
REM 启动 GUI 导表工具的批处理（双击运行）
setlocal

REM 本批处理假定脚本位于仓库的 tools\gui_export_launcher.py，相对于本文件为 ..\tools\gui_export_launcher.py
set SCRIPT_PATH=%~dp0..\tools\gui_export_launcher.py

IF NOT EXIST "%SCRIPT_PATH%" (
    echo GUI 脚本未找到: %SCRIPT_PATH%
    echo 请确认仓库结构未改变，或手动修改本批处理中的脚本路径。
    pause
    exit /b 1
)

REM 优先使用环境中的 python
where python >nul 2>nul
if %ERRORLEVEL%==0 (
    set PYEXEC=python
) else (
    REM 若系统未找到 python，可在此处填写本机 Python 安装路径
    set PYEXEC=C:\Program Files\Python 3.13.0\python.exe
)

echo 使用 Python: %PYEXEC%
REM 检查 python 可执行是否存在（如果配置为 path 名称，则跳过此检查）
if /I "%PYEXEC%" NEQ "python" (
    if not exist "%PYEXEC%" (
        echo 未找到指定的 Python 可执行文件: %PYEXEC%
        echo 请修改批处理或将 Python 添加到 PATH
        pause
        exit /b 1
    )
)

echo 将执行: "%PYEXEC%" "%SCRIPT_PATH%"
"%PYEXEC%" "%SCRIPT_PATH%"

echo.
echo 已完成运行（如有错误，请查看上方输出）。
pause
endlocal
