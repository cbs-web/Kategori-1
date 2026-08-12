@echo off
setlocal
title K-1
cd /d "%~dp0"

if not exist "k1.py" (
    echo k1.py bu klasorde bulunamadi.
    pause
    exit /b 1
)

set "PY_CMD="
python -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 9) else 1)" >nul 2>&1
if not errorlevel 1 set "PY_CMD=python"

if "%PY_CMD%"=="" (
    py -3 -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 9) else 1)" >nul 2>&1
    if not errorlevel 1 set "PY_CMD=py -3"
)

if "%PY_CMD%"=="" (
    echo Uygun Python bulunamadi. K-1 icin Python 3.9 veya daha yeni bir surum gerekir.
    pause
    exit /b 1
)

if exist "bagimlilik_kontrol.py" (
    %PY_CMD% "bagimlilik_kontrol.py"
    if errorlevel 1 (
        echo.
        echo Eksik paketler tamamlanmadan K-1 baslatilamadi.
        pause
        exit /b 1
    )
)

%PY_CMD% "k1.py"
if errorlevel 1 (
    echo.
    echo K-1 baslatilirken hata olustu.
    pause
)
