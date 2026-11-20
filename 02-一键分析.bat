@echo off
chcp 936 >nul
cd /d "%~dp0"
title 02 一键分析

echo ==========================================
echo   02 一键分析
echo ==========================================
echo.

if not exist "星铁装备刷取-自动生成-豆包.xlsx" (
    echo 没有角色表，先自动导出一次...
    python "code\json_关键信息-豆包.py"
    if errorlevel 1 goto fail
    echo.
)

echo 计算中（约 3 秒，别关窗口）...
python "code\main.py" "星铁装备刷取-自动生成-豆包.xlsx" "分析结果-豆包.xlsx"
if errorlevel 1 goto fail

echo.
echo ==========================================
echo   分析完成，打开 分析结果-豆包.xlsx
echo   最实用：角色部位完成度 / 角色权重
echo ==========================================
pause
exit /b 0

:fail
echo.
echo !! 出错，请看上面的报错信息。
pause
exit /b 1
