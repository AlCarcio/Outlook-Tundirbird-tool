@echo off
echo ============================================
echo   MSG to EML Converter
echo ============================================
echo.

:: Try to find Python - prefer py launcher (always available on Windows)
set PYTHON=
set PIP=

:: Option 1: py launcher (most reliable on Windows)
py -3 --version >nul 2>&1
if not errorlevel 1 (
    set PYTHON=py -3
    set PIP=py -3 -m pip
    goto :found
)

:: Option 2: python in PATH
python --version >nul 2>&1
if not errorlevel 1 (
    set PYTHON=python
    set PIP=python -m pip
    goto :found
)

:: Nothing found
echo [ERRORE] Python non trovato!
echo.
echo Scarica Python 3.12 o 3.13 da: https://www.python.org/downloads/
echo IMPORTANTE: durante l'installazione seleziona "Add Python to PATH"
echo.
pause
exit /b 1

:found
:: Show detected version
for /f "tokens=*" %%v in ('%PYTHON% --version 2^>^&1') do echo %%v rilevato - OK
echo.

echo Installazione dipendenze...
%PIP% install -r requirements_msg.txt --quiet
if errorlevel 1 (
    echo [ERRORE] Installazione dipendenze fallita.
    pause
    exit /b 1
)

echo.
echo Avvio MSG to EML Converter...
%PYTHON% main_msg.py
pause
