@echo off
chcp 936 >nul
cd /d "%~dp0"
title 00 安装依赖

echo ==========================================
echo   00 安装 Python 依赖
echo   需要：pandas、openpyxl
echo ==========================================
echo.

rem ---- 1. 先确认有 Python ----
python --version >nul 2>nul
if errorlevel 1 goto nopython

echo 当前 Python：
python --version
echo.

rem ---- 2. 确认有 pip，没有就用 ensurepip 补一个 ----
python -m pip --version >nul 2>nul
if errorlevel 1 goto nopip
goto install

:nopip
echo 没检测到 pip，尝试用 ensurepip 装一个...
python -m ensurepip --upgrade
if errorlevel 1 goto pipfail
echo.

:install
echo [1/3] 安装 pandas ...
python -m pip install --disable-pip-version-check pandas
if errorlevel 1 goto mirror
echo.
echo [2/3] 安装 openpyxl ...
python -m pip install --disable-pip-version-check openpyxl
if errorlevel 1 goto mirror
goto verify

:mirror
echo.
echo 直连官方源没装上，改用清华镜像重试...
python -m pip install --disable-pip-version-check -i https://pypi.tuna.tsinghua.edu.cn/simple pandas openpyxl
if errorlevel 1 goto fail

:verify
echo.
echo [3/3] 验证能否正常 import ...
python -c "import pandas, openpyxl; print('  pandas  ' + pandas.__version__ + '  /  openpyxl  ' + openpyxl.__version__)"
if errorlevel 1 goto fail

echo.
echo ==========================================
echo   装好了，下一步双击 01-一键导出.bat
echo ==========================================
pause
exit /b 0

:nopython
echo [错误] 没找到 python 命令。
echo.
echo 说明 Python 没装，或者装的时候没勾选 Add to PATH。
echo 重新装一遍 Python 3 时，记得勾上 "Add python.exe to PATH"，
echo 装完关掉再重开这个窗口。
echo.
pause
exit /b 1

:pipfail
echo [错误] pip 补装失败，多半是网络问题。
echo 可以先手动装个 Python 3.8 以上版本（自带 pip），再跑这个脚本。
echo.
pause
exit /b 1

:fail
echo.
echo !! 装库失败了，看上面的报错。
echo    网络问题就挂个梯子，或者手动执行：
echo    python -m pip install -i https://pypi.tuna.tsinghua.edu.cn/simple pandas openpyxl
echo.
pause
exit /b 1