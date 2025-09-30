# Archivo: run-kms.ps1
# Launcher interactivo en PowerShell

# URL del script remoto
$URL = 'https://raw.githubusercontent.com/sinister1977/scripts/main/KMS_Suite.v8.3.EN.bat'

# Carpeta temporal donde se guardará el script
$TEMP_DIR = Join-Path $env:LOCALAPPDATA 'Temp\KMS_SuiteLauncher'
if (-not (Test-Path $TEMP_DIR)) {
    New-Item -ItemType Directory -Path $TEMP_DIR -Force | Out-Null
}

$SCRIPT_PATH = Join-Path $TEMP_DIR 'KMS_Suite.v8.3.EN.bat'

# Descargar el script remoto usando curl si está disponible, si no Invoke-WebRequest
Write-Host "Descargando script desde GitHub..."
if (Get-Command curl.exe -ErrorAction SilentlyContinue) {
    curl.exe -L $URL -o $SCRIPT_PATH
} else {
    try {
        Invoke-WebRequest -Uri $URL -OutFile $SCRIPT_PATH -UseBasicParsing -ErrorAction Stop
    } catch {
        Write-Host "Error al descargar el script: $_"
        exit
    }
}

if (-not (Test-Path $SCRIPT_PATH)) {
    Write-Host "Error al descargar el script."
    Pause
    exit
}

Write-Host "Descargado en: $SCRIPT_PATH`n"

# Menú interactivo
do {
    Write-Host "Seleccione una opción:"
    Write-Host "1) Ejecutar normalmente"
    Write-Host "2) Ejecutar como administrador"
    Write-Host "3) Salir"
    $opcion = Read-Host "Ingrese su elección (1-3)"

    switch ($opcion) {
        "1" {
            Write-Host "`nEjecutando normalmente..."
            & "$SCRIPT_PATH"
        }
        "2" {
            Write-Host "`nEjecutando como administrador..."
            Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$SCRIPT_PATH`"" -Verb RunAs -Wait
        }
        "3" {
            Write-Host "Saliendo..."
        }
        default {
            Write-Host "Opción inválida, intente de nuevo.`n"
        }
    }
} while ($opcion -ne "3")
