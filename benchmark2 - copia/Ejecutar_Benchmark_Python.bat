@echo off
REM Lanzador del benchmark en Python usando el Python portable local (sin instalar nada)
setlocal
set "PYDIR=%~dp0.pyportable"
set "PYTHONIOENCODING=utf-8"
set "PYTHONUTF8=1"
cd /d "%~dp0"
"%PYDIR%\python.exe" "%~dp0Benchmark.py"
echo.
echo === Terminado. Excel en: %~dp0Analisis_Energia_CNMC.xlsx ===
pause
