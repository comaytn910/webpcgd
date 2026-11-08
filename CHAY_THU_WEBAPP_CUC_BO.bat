@echo off
chcp 65001 > nul
title HE THONG TRA CUU PCGD - XMC THAI NGUYEN - WEBAPP VIEWER
cd /d "%~dp0"

echo ===============================================================================
echo     SO GIAO DUC VA DAO TAO TINH THAI NGUYEN
echo     HE THONG TRA CUU SO LIEU PHO CAP GIAO DUC - XOA MU CHU 2026
echo     Phien ban Web App Doc lap - Viewer Mode
echo ===============================================================================
echo.

set "PY_CMD="

if exist "%~dp0..\.venv\Scripts\python.exe" (
    set "PY_CMD=%~dp0..\.venv\Scripts\python.exe"
    goto RUN_APP
)

if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" (
    set "PY_CMD=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
    goto RUN_APP
)

where python >nul 2>nul
if %errorlevel% equ 0 (
    set "PY_CMD=python"
    goto RUN_APP
)

echo [LOI] Khong tim thay Python tren may tinh!
pause
exit /b 1

:RUN_APP
echo [*] Dang su dung moi truong: %PY_CMD%
echo [*] Dang khoi chay ung dung Webapp tai cong 8502...
echo [*] Trinh duyet se tu dong mo trong giay lat...
echo.

"%PY_CMD%" -m streamlit run app.py --server.port 8502 --server.headless false

pause
