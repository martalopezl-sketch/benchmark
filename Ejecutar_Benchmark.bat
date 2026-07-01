@echo off
REM Lanzador de doble clic para Benchmark.ps1 (no necesita Python)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Benchmark.ps1"
echo.
pause
