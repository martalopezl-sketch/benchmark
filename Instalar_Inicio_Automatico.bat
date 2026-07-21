@echo off
REM Copia el script al inicio de Windows (sin necesidad de permisos admin)

set "ORIGEN=%~dp0Auto_Benchmark_Lunes.bat"
set "DESTINO=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\Auto_Benchmark_Lunes.bat"

echo Instalando inicio automatico...
echo.
echo Origen: %ORIGEN%
echo Destino: %DESTINO%
echo.

copy /Y "%ORIGEN%" "%DESTINO%"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo === Instalacion exitosa ===
    echo.
    echo El benchmark se ejecutara automaticamente:
    echo - Cada vez que inicies sesion en Windows
    echo - Solo los lunes
    echo - A las 10:00 AM (o cuando enciendas si es despues de las 10)
    echo - Una sola vez por dia
    echo.
    echo Para desinstalar, simplemente borra:
    echo %DESTINO%
) else (
    echo.
    echo ERROR: No se pudo copiar el archivo
)

echo.
pause
