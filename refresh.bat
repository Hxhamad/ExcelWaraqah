@echo off
chcp 65001 >nul
rem Run from this script's own directory (portable)
cd /d "%~dp0"

rem Prefer the Python launcher, fall back to python on PATH
set PY=python
where py >nul 2>nul && set PY=py -3

rem Optional local yfinance upgrade shim (only used if present)
if exist "%LOCALAPPDATA%\Temp\yf_pkg" set PYTHONPATH=%LOCALAPPDATA%\Temp\yf_pkg

if "%1"=="quick" goto quick
if "%1"=="full" goto full
echo 1) quick  (تحديث المحفظة فقط، ~2 دقيقة)
echo 2) full   (تحديث السوق كامل، ~20-40 دقيقة)
set /p c=اختر 1 أو 2:
if "%c%"=="1" goto quick
goto full
:quick
"%PY%" refresh_quick.py
goto done
:full
"%PY%" refresh_full.py
:done
pause
