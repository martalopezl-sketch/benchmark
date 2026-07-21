@echo off
REM Script que se ejecuta al iniciar Windows y verifica si es lunes para correr el benchmark
setlocal enabledelayedexpansion

REM Obtener el día de la semana (1=Lunes, 7=Domingo)
for /f "tokens=1" %%a in ('powershell -Command "(Get-Date).DayOfWeek.value__"') do set DIA=%%a

REM Verificar si es lunes (1)
if not "%DIA%"=="1" (
    exit /b
)

REM Verificar si ya se ejecutó hoy
set "MARCA=%~dp0.ultima_ejecucion.txt"
set "HOY=%date:~-4%%date:~3,2%%date:~0,2%"

if exist "%MARCA%" (
    set /p ULTIMA=<"%MARCA%"
    if "!ULTIMA!"=="%HOY%" (
        exit /b
    )
)

REM Esperar a las 10:00 AM si es antes de esa hora
for /f "tokens=1,2 delims=:" %%a in ("%time%") do (
    set HORA=%%a
    set MIN=%%b
)
set /a "MINUTOS_ACTUALES=%HORA:~-2%*60+%MIN:~0,2%"
set /a "MINUTOS_OBJETIVO=10*60"

if %MINUTOS_ACTUALES% LSS %MINUTOS_OBJETIVO% (
    REM Calcular cuántos segundos esperar
    set /a "ESPERA=(%MINUTOS_OBJETIVO%-%MINUTOS_ACTUALES%)*60"
    timeout /t !ESPERA! /nobreak >nul
)

REM Ejecutar el benchmark
cd /d "%~dp0"
call "%~dp0Ejecutar_Benchmark_Python.bat"

REM Marcar como ejecutado hoy
echo %HOY%>"%MARCA%"
