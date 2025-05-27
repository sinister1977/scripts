# Obtener información del sistema operativo
$osInfo = Get-CimInstance Win32_OperatingSystem
$strOsInfo = @"
Información del sistema operativo:

Sistema Operativo: $($osInfo.Caption) $($osInfo.Version) (Build $($osInfo.BuildNumber))
Arquitectura: $($osInfo.OSArchitecture)
"@

# Obtener información del sistema
$objWMIService = Get-CimInstance -ClassName Win32_ComputerSystem

# Obtener direcciones MAC de tarjetas de red
$colNetworkAdapters = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration
$strMacAddresses = "`r`nDirecciones MAC:`r`n`r`n"
foreach ($objAdapter in $colNetworkAdapters) {
    if ($objAdapter.MACAddress -ne $null) {
        $strMacAddresses += "Descripción: $($objAdapter.Description)`r`n"
        $strMacAddresses += "Dirección MAC: $($objAdapter.MACAddress)`r`n`r`n"
    }
}

# Obtener información de la computadora
$colBIOS = Get-CimInstance -Class Win32_BIOS
$colPhysicalMemory = Get-CimInstance -Class Win32_PhysicalMemory
$colDiskDrive = Get-CimInstance -Class Win32_DiskDrive

# Obtener el Service Tag (Número de serie del equipo)
$serviceTag = (Get-CimInstance Win32_BIOS | Select-Object -ExpandProperty SerialNumber).Trim()

# Obtener sesión de usuario
$userSession = whoami

# Obtener la fecha y hora actual en formato DDMMYYYY-HHMMSS
$timestamp = Get-Date -Format "ddMMyyyy-HHmmss"

# Obtener direcciones IP internas y pública
$networkInterfaces = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' }
$internalIPs = $networkInterfaces | ForEach-Object {
    $interfaceName = $_.Name
    $interfaceIP = (Get-NetIPAddress -InterfaceAlias $interfaceName | Where-Object { $_.AddressFamily -eq 'IPv4' }).IPAddress
    "Interfaz: $interfaceName - IP Interna: $interfaceIP"
} | Out-String

$publicIP = (Invoke-WebRequest -Uri "http://ipinfo.io/ip").Content.Trim()
$locationData = Invoke-WebRequest -Uri "http://ipinfo.io/$publicIP/json" | ConvertFrom-Json
$location = "$($locationData.city), $($locationData.country)"
$provider = $locationData.org

# Obtener información del disco duro
$strDiskInfo = "`r`nCapacidad total del disco duro instalado y tipo de disco duro:`r`n" 
foreach ($disk in $colDiskDrive) {
    $strDiskInfo += "Modelo: $($disk.Model)`r`n"
    $strDiskInfo += "Capacidad total: $([math]::Round($disk.Size / 1GB, 2)) GB`r`n"
    $strDiskInfo += "Tipo: $($disk.MediaType)`r`n"
}

# Obtener información de la memoria RAM
$strRAMInfo = "`r`nInformación de la memoria RAM instalada:`r`n" 
foreach ($ram in $colPhysicalMemory) {
    $strRAMInfo += "Fabricante: $($ram.Manufacturer)`r`n"
    $strRAMInfo += "Capacidad: $([math]::Round($ram.Capacity / 1GB, 2)) GB`r`n"
    $strRAMInfo += "Tipo: $($ram.MemoryType)`r`n"  
}

# Obtener lista de programas instalados con fecha de instalación
$installedPrograms = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | 
    Select-Object DisplayName, InstallDate | Sort-Object DisplayName | Format-Table -AutoSize | Out-String

# Obtener información del antivirus
$antivirusInfo = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntivirusProduct | 
    Select-Object displayName, productState | Format-Table -AutoSize | Out-String

# Construir la cadena de información
$strInfo = "$strOsInfo`r`n"
$strInfo += "Nombre del equipo: $($objWMIService.Name)`r`n"
$strInfo += "Fabricante: $($objWMIService.Manufacturer)`r`n"
$strInfo += "Modelo: $($objWMIService.Model)`r`n"
$strInfo += "Service Tag: $serviceTag`r`n"
$strInfo += "Usuario que inició sesión: $userSession`r`n"
$strInfo += "Direcciones IP internas:`r`n$internalIPs`r`n"
$strInfo += "Dirección IP pública: $publicIP`r`nUbicación: $location`r`nProveedor: $provider`r`n"
$strInfo += "$strMacAddresses"
$strInfo += "$strDiskInfo"
$strInfo += "$strRAMInfo"
$strInfo += "`r`nProgramas instalados:`r`n$installedPrograms"
$strInfo += "`r`nInformación del antivirus:`r`n$antivirusInfo"

# Guardar información en archivo con nombre personalizado
$strFilePath = "$env:USERPROFILE\Inventario_$($objWMIService.Name)_$timestamp.txt"

# Verificar si el archivo está en uso antes de escribir
$attempts = 5
while ($attempts -gt 0) {
    try {
        $strInfo | Out-File -FilePath $strFilePath -Encoding UTF8 -Force -Confirm:$false
        Write-Host "Archivo guardado exitosamente."
        break
    } catch {
        Write-Host "El archivo está en uso, reintentando..."
        Start-Sleep -Seconds 2
        $attempts--
    }
}

# Enviar correo electrónico con el archivo adjunto
$smtp_server = "smtp.gmail.com"
$smtp_port = 587
$smtp_username = "solexactivosscl@gmail.com"
$smtp_password = "fznm abuv tsjl nbcj"
$cred = New-Object System.Net.NetworkCredential($smtp_username, $smtp_password)
$email_to = @("sergio.sanmartin@solex.biz", "ernesto.perez@solex.biz", "richard.buitrago@solex.biz")

$MailMessage = New-Object System.Net.Mail.MailMessage
$MailMessage.From = $smtp_username
foreach ($recipient in $email_to) {
    $MailMessage.To.Add($recipient)
}
$MailMessage.Subject = "Inventario SOLEX - $($objWMIService.Name) - Service Tag: $serviceTag - Usuario: $userSession"
$MailMessage.Body = "Adjunto encontrarás el inventario del equipo."
$MailMessage.IsBodyHtml = $false
$MailMessage.Attachments.Add((New-Object System.Net.Mail.Attachment($strFilePath)))

$SmtpClient = New-Object System.Net.Mail.SmtpClient($smtp_server, $smtp_port)
$SmtpClient.Credentials = $cred
$SmtpClient.EnableSsl = $true
$SmtpClient.Send($MailMessage)

Write-Host "Correo enviado exitosamente."
