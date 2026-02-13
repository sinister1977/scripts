# Obtener la fecha y hora actual en formato DDMMYYYY-HHMMSS
$timestamp = Get-Date -Format "ddMMyyyy-HHmmss"

# Obtener información del sistema operativo
$osInfo = Get-CimInstance Win32_OperatingSystem

# Obtener información del procesador
$cpu = Get-CimInstance Win32_Processor
$cpuName = $cpu.Name.Trim()
$manufacturer = $cpu.Manufacturer
$architecture = switch ($cpu.Architecture) {
    0 { "x86" }
    1 { "MIPS" }
    2 { "Alpha" }
    3 { "PowerPC" }
    5 { "ARM" }
    6 { "Itanium" }
    9 { "x64" }
    default { "Desconocida" }
}
$generation = "N/A"
if ($manufacturer -match "Intel") {
    if ($cpuName -match "i[0-9]-([0-9]{3,4})[A-Za-z]?") {
        $genNumber = $Matches[1].Substring(0, 1)
        $generation = "${genNumber}ª Generación"
    } elseif ($cpuName -match "Intel.*(Core|Pentium|Celeron)") {
        $generation = "Modelo antiguo (pre-generaciones)"
    }
} elseif ($manufacturer -match "AMD") {
    $generation = "AMD " + ($cpuName -replace ".*(Ryzen|EPYC|Threadripper|Athlon|FX).*", '$1')
}

# Obtener información del sistema
$objWMIService = Get-CimInstance -ClassName Win32_ComputerSystem
$colBIOS = Get-CimInstance -Class Win32_BIOS
$serviceTag = ($colBIOS | Select-Object -ExpandProperty SerialNumber).Trim()

# Obtener sesión de usuario
$userSession = whoami

# Obtener direcciones IP internas
$networkInterfaces = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' }
$internalIPs = $networkInterfaces | ForEach-Object {
    $interfaceName = $_.Name
    $interfaceIP = (Get-NetIPAddress -InterfaceAlias $interfaceName | Where-Object { $_.AddressFamily -eq 'IPv4' }).IPAddress
    "Interfaz: $interfaceName - IP Interna: $interfaceIP"
} | Out-String

# Obtener IP pública y localización
$publicIP = (Invoke-WebRequest -Uri "http://ipinfo.io/ip").Content.Trim()
$locationData = Invoke-WebRequest -Uri "http://ipinfo.io/$publicIP/json" | ConvertFrom-Json
$location = "$($locationData.city), $($locationData.country)"
$provider = $locationData.org

# Obtener direcciones MAC
$colNetworkAdapters = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration
$macAddresses = foreach ($objAdapter in $colNetworkAdapters) {
    if ($objAdapter.MACAddress -ne $null) {
        "<li><b>Descripción:</b> $($objAdapter.Description) - <b>MAC:</b> $($objAdapter.MACAddress)</li>"
    }
}

# Obtener discos duros
$colDiskDrive = Get-CimInstance -Class Win32_DiskDrive
$diskInfo = foreach ($disk in $colDiskDrive) {
    "<li><b>Modelo:</b> $($disk.Model) - <b>Capacidad:</b> $([math]::Round($disk.Size / 1GB, 2)) GB - <b>Tipo:</b> $($disk.MediaType)</li>"
}

# RAM
$colPhysicalMemory = Get-CimInstance -Class Win32_PhysicalMemory
$ramInfo = foreach ($ram in $colPhysicalMemory) {
    "<li><b>Fabricante:</b> $($ram.Manufacturer) - <b>Capacidad:</b> $([math]::Round($ram.Capacity / 1GB, 2)) GB - <b>Tipo:</b> $($ram.MemoryType)</li>"
}

# Programas instalados
$installedPrograms = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* |
    Select-Object DisplayName, InstallDate | Sort-Object DisplayName |
    Where-Object { $_.DisplayName } |
    ForEach-Object {
        "<tr><td>$($_.DisplayName)</td><td>$($_.InstallDate)</td></tr>"
    }

# Antivirus
$antivirusInfo = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntivirusProduct |
    ForEach-Object {
        "<tr><td>$($_.displayName)</td><td>$($_.productState)</td></tr>"
    }

# Crear HTML
$html = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset='UTF-8'>
    <title>Inventario del equipo $($objWMIService.Name)</title>
    <style>
        body { font-family: Arial, sans-serif; }
        h2 { color: #2E6E9E; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
        th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        ul { padding-left: 20px; }
    </style>
</head>
<body>
    <h2>Inventario del equipo - $($objWMIService.Name)</h2>
    <p><b>Fecha:</b> $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")</p>
    <h3>Sistema Operativo</h3>
    <p>$($osInfo.Caption) $($osInfo.Version) (Build $($osInfo.BuildNumber))<br>Arquitectura: $($osInfo.OSArchitecture)</p>
    <h3>Procesador</h3>
    <ul>
        <li><b>Fabricante:</b> $($manufacturer -replace 'Genuine', '')</li>
        <li><b>Modelo:</b> $cpuName</li>
        <li><b>Generación:</b> $generation</li>
        <li><b>Arquitectura:</b> $architecture</li>
        <li><b>Núcleos:</b> $($cpu.NumberOfCores) físicos, $($cpu.NumberOfLogicalProcessors) lógicos</li>
    </ul>
    <h3>Equipo</h3>
    <ul>
        <li><b>Nombre:</b> $($objWMIService.Name)</li>
        <li><b>Fabricante:</b> $($objWMIService.Manufacturer)</li>
        <li><b>Modelo:</b> $($objWMIService.Model)</li>
        <li><b>Service Tag:</b> $serviceTag</li>
        <li><b>Usuario actual:</b> $userSession</li>
    </ul>
    <h3>Red</h3>
    <p><b>IP pública:</b> $publicIP<br><b>Ubicación:</b> $location<br><b>Proveedor:</b> $provider</p>
    <p><b>IPs internas:</b><br><pre>$internalIPs</pre></p>
    <ul>$($macAddresses -join "`n")</ul>
    <h3>Discos</h3>
    <ul>$($diskInfo -join "`n")</ul>
    <h3>RAM</h3>
    <ul>$($ramInfo -join "`n")</ul>
    <h3>Programas instalados</h3>
    <table><tr><th>Nombre</th><th>Fecha de instalación</th></tr>
    $($installedPrograms -join "`n")
    </table>
    <h3>Antivirus</h3>
    <table><tr><th>Nombre</th><th>Estado</th></tr>
    $($antivirusInfo -join "`n")
    </table>
</body>
</html>
"@

# Guardar HTML
$htmlPath = "$env:USERPROFILE\Inventario_$($objWMIService.Name)_$timestamp.html"
$html | Out-File -FilePath $htmlPath -Encoding UTF8 -Force

# Enviar correo
$smtpServer = (-join [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('c210cC1yZWxheS5icmV2by5jb20=')).ToCharArray() | ForEach-Object {$_}) -join ''
$smtpPort = 587
$smtpUser = (-join [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('YTIyNGFhMDAxQHNtdHAtYnJldm8uY29t')).ToCharArray() | ForEach-Object {$_}) -join ''
$smtpPassword = (-join [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('eHNtdHBzaWItOGIyYTRmODA2ZDQ0NzJiYmQ4M2RhYjIzNTFiYmMyNTIxN2I2OWYxNmNiYmMzMTRmYjZlZjViNzY0YWZhYjBhMC1scEgyTXJlU1hlaEw1OVRY')).ToCharArray() | ForEach-Object {$_}) -join ''
$emailFrom = (-join [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('c29sZXhhY3Rpdm9zc2NsQGdtYWlsLmNvbQ==')).ToCharArray() | ForEach-Object {$_}) -join ''
$emailTo = (-join [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('c2VyZ2lvLnNhbm1hcnRpbkBzb2xleC5iaXo=')).ToCharArray() | ForEach-Object {$_}) -join ''

$MailMessage = New-Object System.Net.Mail.MailMessage
$MailMessage.From = $smtp_username
foreach ($recipient in $email_to) {
    $MailMessage.To.Add($recipient)
}
$MailMessage.Subject = "Inventario SOLEX - $($objWMIService.Name) - Service Tag: $serviceTag - Usuario: $userSession"
$MailMessage.Body = "<p>Adjunto encontrarás el inventario del equipo <b>$($objWMIService.Name)</b>.</p>"
$MailMessage.IsBodyHtml = $true
$MailMessage.Attachments.Add((New-Object System.Net.Mail.Attachment($htmlPath)))

$SmtpClient = New-Object System.Net.Mail.SmtpClient($smtp_server, $smtp_port)
$SmtpClient.Credentials = $cred
$SmtpClient.EnableSsl = $true
$SmtpClient.Send($MailMessage)

Write-Host "Correo enviado exitosamente con archivo HTML adjunto."

