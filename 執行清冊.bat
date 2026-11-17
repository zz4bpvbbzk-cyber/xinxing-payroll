@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
title 新興國小 薪資清冊生成器
cd /d "%~dp0"

echo ══════════════════════════════════════════════
echo   新興國小 薪資清冊生成器(雙擊版)
echo   固定節數 ‧ 超鐘點 ‧ 直播共學 ‧ 越南語 ‧ 短代
echo ══════════════════════════════════════════════
echo.

REM ---- 0) Python 檢查 ----
set PY=python
python --version >nul 2>&1
if errorlevel 1 (
    where py >nul 2>&1 && set PY=py || (
    echo [錯誤] 找不到 Python。請先安裝:
    echo        https://www.microsoft.com/store/productId/9NCVDN91XZQP ^(微軟商店版即可^)
    echo        或到 https://www.python.org/downloads/ 下載,安裝時勾選 "Add Python to PATH"
    pause
    exit /b 1
    )
)
echo [0] Python 檢查通過 !PY!
echo.

REM ---- 1) openpyxl 檢查/安裝 ----
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

REM ---- 2) 選學期(掃描 config 資料夾) ----
:choose
echo [2] 選擇學期設定檔:
echo.
set N=0
set "CFG1=" & set "CFG2=" & set "CFG3=" & set "CFG4=" & set "CFG5=" & set "CFG6=" & set "CFG7=" & set "CFG8=" & set "CFG9="
for %%f in (config\*.xlsx config-*\*.xlsx) do (
    set /a N+=1
    set "CFG!N!=%%f"
    echo     !N!. %%~nf
)
if %N%==0 (
    echo.
    echo [錯誤] config 資料夾裡沒有任何設定檔 .xlsx
    echo        請先聯絡教學組長建立,或到 GitHub 下說明建立新學期設定檔
    pause
    exit /b 1
)
echo.
choice /c 123456789 /n /m "    請按數字選擇 (1-%N%),按 Esc 離開: "
if errorlevel 9 set "SEL=%CFG9%" & goto run
if errorlevel 8 set "SEL=%CFG8%" & goto run
if errorlevel 7 set "SEL=%CFG7%" & goto run
if errorlevel 6 set "SEL=%CFG6%" & goto run
if errorlevel 5 set "SEL=%CFG5%" & goto run
if errorlevel 4 set "SEL=%CFG4%" & goto run
if errorlevel 3 set "SEL=%CFG3%" & goto run
if errorlevel 2 set "SEL=%CFG2%" & goto run
if errorlevel 1 set "SEL=%CFG1%" & goto run
goto choose

:run
echo.
echo [3] 開始生成:%SEL%
echo     ────────────────────────────────────────────
!PY! scripts\gen.py --config "%SEL%"
if errorlevel 1 (
    echo     ────────────────────────────────────────────
    echo [錯誤] 生成失敗!上方訊息指出問題(驗證沒過或設定有誤)。
    echo        請把整個黑色視窗的內容拍照/複製,連絡教學組長或開 Genspark 對話處理。
    pause
    exit /b 1
)
echo     ────────────────────────────────────────────
echo [4] 完成!

REM ---- 5) 找出輸出資料夾並自動開啟 ----
for /f "delims=" %%d in ('dir /b /ad out 2^>nul ^| sort /r') do (
    explorer "%~dp0out\%%d"
    goto :end
)

:end
echo.
echo 輸出資料夾已開啟(裡面有各月清冊與驗算報告)。
echo 這個視窗可以關閉了。
pause
