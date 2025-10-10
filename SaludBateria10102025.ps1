# ================================================ 
# BatteryHealthReport.ps1
# Genera un informe claro con la salud de la bateria y lo envia por correo
# ================================================

# Directorio donde se guardara el reporte
$reportDir = "C:\Temp"
if (-not (Test-Path $reportDir)) { New-Item -ItemType Directory -Path $reportDir | Out-Null }

$timestamp = Get-Date -Format 'yyyyMMdd_HHmm'
$htmlReport = "$reportDir\battery-report_$timestamp.html"
$textReport = "$reportDir\battery-health_$timestamp.txt"

# Generar informe oficial de Windows
powercfg /batteryreport /output $htmlReport | Out-Null

# Leer HTML
$rt = Get-Content $htmlReport -Raw -Encoding UTF8
$rt = $rt -replace '&nbsp;', ' '

function CleanNum([string]$s){
    if (-not $s) { return 0 }
    $digits = [regex]::Replace($s, '[^\d]', '')
    if ($digits -eq '') { return 0 }
    return [int]$digits
}

# Función para eliminar acentos
function Remove-Accents([string]$text){
    $text = $text -replace '[áÁ]', 'a'
    $text = $text -replace '[éÉ]', 'e'
    $text = $text -replace '[íÍ]', 'i'
    $text = $text -replace '[óÓ]', 'o'
    $text = $text -replace '[úÚ]', 'u'
    $text = $text -replace '[ñÑ]', 'n'
    return $text
}

# Variables iniciales
$designCapacity = 0
$fullChargeCapacity = 0
$unit = "mWh"

# Patrones para ingles y español
$patterns = @(
    @{name='design'; rx='(?i)Design Capacity<\/td>\s*<td[^>]*>\s*([0-9\.,\s]+)\s*(mWh|mAh)?'},
    @{name='full';   rx='(?i)Full Charge Capacity<\/td>\s*<td[^>]*>\s*([0-9\.,\s]+)\s*(mWh|mAh)?'},
    @{name='design'; rx='(?i)Capacidad de dise[oó]o<\/td>\s*<td[^>]*>\s*([0-9\.,\s]+)\s*(mWh|mAh)?'},
    @{name='full';   rx='(?i)Capacidad (?:de )?carga completa<\/td>\s*<td[^>]*>\s*([0-9\.,\s]+)\s*(mWh|mAh)?'}
)

# Buscar valores
foreach ($p in $patterns) {
    $m = [regex]::Match($rt, $p.rx)
    if ($m.Success) {
        $val = CleanNum($m.Groups[1].Value)
        switch ($p.name) {
            'design' { if ($val -gt 0) { $designCapacity = $val }; if ($m.Groups.Count -ge 3 -and $m.Groups[2].Value) { $unit = $m.Groups[2].Value } }
            'full'   { if ($val -gt 0) { $fullChargeCapacity = $val }; if ($m.Groups.Count -ge 3 -and $m.Groups[2].Value) { $unit = $m.Groups[2].Value } }
        }
    }
}

# Fallback
if ($designCapacity -eq 0 -or $fullChargeCapacity -eq 0) {
    $markers = @('Installed batteries','Baterias instaladas')
    $pos = -1
    foreach ($marker in $markers) {
        $pos = $rt.IndexOf($marker, [StringComparison]::OrdinalIgnoreCase)
        if ($pos -ge 0) { break }
    }
    if ($pos -ge 0) {
        $length = [Math]::Min(1500, $rt.Length - $pos)
        $sec = $rt.Substring($pos, $length)
        $matches = [regex]::Matches($sec, '([0-9][0-9\.,\s]+)\s*(mWh|mAh)', 'IgnoreCase')
        $nums = @()
        foreach ($mm in $matches) {
            $n = CleanNum($mm.Groups[1].Value)
            if ($n -gt 0) { $nums += $n }
        }
        $nums = $nums | Sort-Object -Descending -Unique
        if ($nums.Count -ge 2) {
            $designCapacity = $nums[0]
            $fullChargeCapacity = $nums[1]
        }
    }
}

# Calcular salud
if ($designCapacity -gt 0 -and $fullChargeCapacity -gt 0) {
    $batteryHealth = [math]::Round(($fullChargeCapacity / $designCapacity) * 100, 1)
} else {
    $batteryHealth = 'No disponible'
}

# Duracion total de bateria nueva (horas)
$duracionTotalHoras = 2  

# Duracion aproximada en minutos
if ($batteryHealth -ne 'No disponible') {
    $duracionActualMin = [math]::Round($duracionTotalHoras * 60 * ($batteryHealth / 100))
    $duracionTexto = "$duracionActualMin min aprox."
} else {
    $duracionTexto = "No disponible"
}

# Numero de serie
$serialNumber = (wmic bios get serialnumber | Select-String -Pattern "\S+" | Select-Object -Last 1).ToString().Trim()

# Usuario actual
$currentUser = (whoami).Trim()

# Crear resumen
$summary = @"
==========================================
 INFORME DE SALUD DE BATERIA - $env:COMPUTERNAME
 Numero de serie:               $serialNumber
 Usuario actual:                $currentUser
 Fecha: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
==========================================

Capacidad de diseno (fabrica):  $designCapacity $unit
Capacidad maxima actual:        $fullChargeCapacity $unit
Salud estimada de bateria:      $batteryHealth %
Duracion aproximada actual:     $duracionTexto

Archivo HTML completo: $htmlReport
==========================================
"@

# Eliminar acentos para el resumen
$summary = Remove-Accents $summary

# Guardar resultado con UTF8 sin BOM
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($textReport, $summary, $utf8NoBom)

# Mostrar en pantalla
Write-Host $summary

# -------------------------------------------
# Configuracion correo
$smtp_server = "smtp.gmail.com"
$smtp_port = 587
$smtp_username = "solexactivosscl@gmail.com"
$smtp_password = "fznm abuv tsjl nbcj"
$email_to = @("sergio.sanmartin@solex.biz")

# Crear PSCredential
$securePassword = ConvertTo-SecureString $smtp_password -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential($smtp_username, $securePassword)

# Enviar correo
Send-MailMessage `
    -From $smtp_username `
    -To $email_to `
    -Subject "Informe de salud de bateria - $env:COMPUTERNAME" `
    -Body $summary `
    -SmtpServer $smtp_server `
    -Port $smtp_port `
    -Credential $cred `
    -UseSsl
