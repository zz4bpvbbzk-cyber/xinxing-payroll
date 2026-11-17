@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
title 新興國小 薪資清冊生成器
cd /d "%~dp0"

echo ============================================================
echo   新興國小 薪資清冊生成器
echo   固定節數 · 超鐘點 · 直播共學 · 越南語 · 短代
echo ============================================================
echo.

set PY=python
python --version >nul 2>&1
if errorlevel 1 (
    where py >nul 2>&1 && set PY=py || (
    echo [錯誤] 找不到 Python,請先安裝:
    echo        https://www.python.org/downloads/
    echo        安裝時務必勾選 Add python.exe to PATH
    pause
    exit /b 1
    )
)
echo [0] Python 檢查通過
!PY! -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo [1] 缺少 openpyxl 套件,安裝中...
    !PY! -m pip install openpyxl -q
    !PY! -c "import openpyxl" >nul 2>&1 || (
        echo [錯誤] openpyxl 安裝失敗,請檢查網路後重試
        pause
        exit /b 1
    )
)
echo [1] openpyxl 就緒
echo.

:choose
echo [2] 選擇學期設定檔:
echo.
set N=0
for /f "delims=" %%f in ('dir /s /b "config-*.xlsx" 2^>nul') do (
    set /a N+=1
    set "CFG!N!=%%f"
    echo     !N!. %%~nxf   〔%%~dpf〕
)
if %N%==0 goto noloc
if %N%==1 set "SEL=%CFG1%" & goto run
echo.
choice /c 123456789 /n /m "    請按數字選擇: "
if errorlevel 9 set "SEL=%CFG9%" & goto run
if errorlevel 8 set "SEL=%CFG8%" & goto run
if errorlevel 7 set "SEL=%CFG7%" & goto run
if errorlevel 6 set "SEL=%CFG6%" & goto run
if errorlevel 5 set "SEL=%CFG5%" & goto run
if errorlevel 4 set "SEL=%CFG4%" & goto run
if errorlevel 3 set "SEL=%CFG3%" & goto run
if errorlevel 2 set "SEL=%CFG2%" & goto run
set "SEL=%CFG1%" & goto run

:noloc
echo.
echo [錯誤] 本資料夾及子資料夾中找不到任何 config-XXXX-X.xlsx 學期設定檔。
echo        本程式位置:%~dp0
echo        正確結構:bat 與 scripts、config-115-1、templates 同一層
echo        〔從私有倉庫 Code → Download ZIP 解壓即是〕
pause
exit /b 1

:run
echo.
echo [3] 開始生成:%SEL%
echo     ────────────────────────────────────────────
!PY! scripts\gen.py --config "%SEL%"
if errorlevel 1 (
    echo     ────────────────────────────────────────────
    echo [錯誤] 生成失敗!上方訊息指出問題,請拍照連絡教學組長處理。
    pause
    exit /b 1
)
echo     ────────────────────────────────────────────
echo [4] 完成!

for /f "delims=" %%d in ('dir /b /ad out 2^>nul ^| sort /r') do (
    start "" "%~dp0out\%%d"
    goto :end
)

:end
echo.
echo 輸出資料夾已開啟,裡面有各月清冊與驗算報告。本視窗可關閉。
pause
