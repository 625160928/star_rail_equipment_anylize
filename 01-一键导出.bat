@echo off
chcp 936 >nul
cd /d "%~dp0"
title 01 导出数据

echo ==========================================
echo   01 导出数据
echo ==========================================
echo.
echo [1/2] 生成角色表...
python "code\json_关键信息-豆包.py"
if errorlevel 1 goto fail
echo.
echo [2/2] 生成全量记录...
python "code\json_全量提取-豆包.py"
if errorlevel 1 goto fail
echo.
echo ==========================================
echo   导出完成，下一步双击 02-一键分析.bat
echo ==========================================
pause
exit /b 0

:fail
echo.
echo !! 出错，请看上面的报错信息。
echo    常见原因：没装 Python，或没装 pandas / openpyxl
pause
exit /b 1