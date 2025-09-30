@echo off
setlocal enabledelayedexpansion

:: URL del script remoto
set "URL=https://raw.githubusercontent.com/sinister1977/scripts/main/KMS_Suite.v8.3.EN.bat"

:: Carpeta temporal donde se guardará el script
set "TEMP_DIR=%USERPROFILE%\AppData\Local\Temp\KMS_SuiteLauncher"
if not exist "%TEMP_DIR%" mkdir "%TEMP_DIR%"

set "SCRIPT_PATH=%TEMP_DIR%\KMS_Suite.v8.3.EN.bat"

:: Descargar el script remoto usando curl
echo Descargando script desde GitHub...
curl -L "%URL%" -o "%SCRIPT_PATH%"
if not exist "%SCRIPT_PATH%" (
    echo Error al descargar el script.
    pause
    exit /b
)

echo Descargado en: %SCRIPT_PATH%
echo.

:MENU
echo Seleccione una opcion:
echo 1) Ejecutar normalmente
echo 2) Ejecutar como administrador
echo 3) Salir
set /p opcion="Ingrese su eleccion (1-3): "

if "%opcion%"=="1" (
    echo Ejecutando normalmente...
    "%SCRIPT_PATH%"
    goto MENU
) else if "%opcion%"=="2" (
    echo Ejecutando como administrador...
    powershell -NoProfile -Command "Start-Process cmd.exe -ArgumentList '/c ""%SCRIPT_PATH%""' -Verb RunAs -Wait"
    goto MENU
) else if "%opcion%"=="3" (
    echo Saliendo...
    exit /b
) else (
    echo Opcion invalida, intente de nuevo.
    goto MENU
)
