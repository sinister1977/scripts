# Obtener informacion del sistema operativo
$osInfo = Get-CimInstance Win32_OperatingSystem
$strOsInfo = @"
InformaciÃ³n del sistema operativo:

Sistema Operativo: $($osInfo.Caption) $($osInfo.Version) (Build $($osInfo.BuildNumber))
Arquitectura: $($osInfo.OSArchitecture)
"@

# Obtener informacion detallada del procesador
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

# Detectar generacion (para Intel)
$generation = "N/A"
if ($manufacturer -match "Intel") {
    if ($cpuName -match "i[0-9]-([0-9]{3,4})[A-Za-z]?") {
        $genNumber = $Matches[1].Substring(0, 1)
        $generation = "${genNumber}Âª GeneraciÃ³n"
    } elseif ($cpuName -match "Intel.*(Core|Pentium|Celeron)") {
        $generation = "Modelo antiguo (pre-generaciones)"
    }
} elseif ($manufacturer -match "AMD") {
    $generation = "AMD " + ($cpuName -replace ".*(Ryzen|EPYC|Threadripper|Athlon|FX).*", '$1')
}

$strCpuInfo = @"
Informacion del procesador:

Fabricante:   $($manufacturer -replace 'Genuine', '')
Modelo:       $cpuName
GeneraciÃ³n:   $generation
Arquitectura: $architecture
NÃºcleos:      $($cpu.NumberOfCores) fÃ­sicos, $($cpu.NumberOfLogicalProcessors) lÃ³gicos
"@

# Obtener informacion del sistema
$objWMIService = Get-CimInstance -ClassName Win32_ComputerSystem

# Obtener direcciones MAC de tarjetas de red
$colNetworkAdapters = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration
$strMacAddresses = "`r`nDirecciones MAC:`r`n`r`n"
foreach ($objAdapter in $colNetworkAdapters) {
    if ($objAdapter.MACAddress -ne $null) {
        $strMacAddresses += "DescripciÃ³n: $($objAdapter.Description)`r`n"
        $strMacAddresses += "DirecciÃ³n MAC: $($objAdapter.MACAddress)`r`n`r`n"
    }
}

# Obtener informacion de la computadora
$colBIOS = Get-CimInstance -Class Win32_BIOS
$colPhysicalMemory = Get-CimInstance -Class Win32_PhysicalMemory
$colDiskDrive = Get-CimInstance -Class Win32_DiskDrive

# Obtener el Service Tag (Numero de serie del equipo)
$serviceTag = (Get-CimInstance Win32_BIOS | Select-Object -ExpandProperty SerialNumber).Trim()

# Obtener sesion de usuario
$userSession = whoami

# Obtener la fecha y hora actual en formato DDMMYYYY-HHMMSS
$timestamp = Get-Date -Format "ddMMyyyy-HHmmss"

# Obtener direcciones IP internas y publica
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

# Obtener informacion del disco duro
$strDiskInfo = "`r`nCapacidad total del disco duro instalado y tipo de disco duro:`r`n" 
foreach ($disk in $colDiskDrive) {
    $strDiskInfo += "Modelo: $($disk.Model)`r`n"
    $strDiskInfo += "Capacidad total: $([math]::Round($disk.Size / 1GB, 2)) GB`r`n"
    $strDiskInfo += "Tipo: $($disk.MediaType)`r`n"
}

# Obtener informacion de la memoria RAM
$strRAMInfo = "`r`nInformaciÃ³n de la memoria RAM instalada:`r`n" 
foreach ($ram in $colPhysicalMemory) {
    $strRAMInfo += "Fabricante: $($ram.Manufacturer)`r`n"
    $strRAMInfo += "Capacidad: $([math]::Round($ram.Capacity / 1GB, 2)) GB`r`n"
    $strRAMInfo += "Tipo: $($ram.MemoryType)`r`n"  
}

# Obtener lista de programas instalados con fecha de instalaciÃ³n
$installedPrograms = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | 
    Select-Object DisplayName, InstallDate | Sort-Object DisplayName | Format-Table -AutoSize | Out-String

# Obtener informacion del antivirus
$antivirusInfo = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntivirusProduct | 
    Select-Object displayName, productState | Format-Table -AutoSize | Out-String

# Construir la cadena de informacion
$strInfo = "$strOsInfo`r`n"
$strInfo += "$strCpuInfo`r`n"
$strInfo += "Nombre del equipo: $($objWMIService.Name)`r`n"
$strInfo += "Fabricante: $($objWMIService.Manufacturer)`r`n"
$strInfo += "Modelo: $($objWMIService.Model)`r`n"
$strInfo += "Service Tag: $serviceTag`r`n"
$strInfo += "Usuario que iniciÃ³ sesiÃ³n: $userSession`r`n"
$strInfo += "Direcciones IP internas:`r`n$internalIPs`r`n"
$strInfo += "DirecciÃ³n IP pÃºblica: $publicIP`r`nUbicaciÃ³n: $location`r`nProveedor: $provider`r`n"
$strInfo += "$strMacAddresses"
$strInfo += "$strDiskInfo"
$strInfo += "$strRAMInfo"
$strInfo += "`r`nProgramas instalados:`r`n$installedPrograms"
$strInfo += "`r`nInformaciÃ³n del antivirus:`r`n$antivirusInfo"

# Guardar informacion en archivo con nombre personalizado
$strFilePath = "$env:USERPROFILE\Inventario_$($objWMIService.Name)_$timestamp.txt"

# Verificar si el archivo esta en uso antes de escribir
$attempts = 5
while ($attempts -gt 0) {
    try {
        $strInfo | Out-File -FilePath $strFilePath -Encoding UTF8 -Force -Confirm:$false
        Write-Host "Archivo guardado exitosamente."
        break
    } catch {
        Write-Host "El archivo estÃ¡ en uso, reintentando..."
        Start-Sleep -Seconds 2
        $attempts--
    }
}


# COMPLEJO - Reemplaza estas lÃ­neas:
$smtp_server = (-join [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('c210cC1yZWxheS5icmV2by5jb20=')).ToCharArray() | ForEach-Object {$_}) -join ''
$smtp_port = 587
$smtp_username = (-join [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('YTIyNGFhMDAxQHNtdHAtYnJldm8uY29t')).ToCharArray() | ForEach-Object {$_}) -join ''
$smtp_password = (-join [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('eHNtdHBzaWItOGIyYTRmODA2ZDQ0NzJiYmQ4M2RhYjIzNTFiYmMyNTIxN2I2OWYxNmNiYmMzMTRmYjZlZjViNzY0YWZhYjBhMC1LMlI1YjFtd3VadERsRm5G')).ToCharArray() | ForEach-Object {$_}) -join ''
$cred = New-Object System.Net.NetworkCredential($smtp_username, $smtp_password)
$email_to = @((-join [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('dGlja2V0c0Bzb2xleC5iaXo=')).ToCharArray() | ForEach-Object {$_}) -join '')

$MailMessage = New-Object System.Net.Mail.MailMessage
$MailMessage.From = (-join [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('c29sZXhhY3Rpdm9zc2NsQGdtYWlsLmNvbQ==')).ToCharArray() | ForEach-Object {$_}) -join ''


foreach ($recipient in $email_to) {
    $MailMessage.To.Add($recipient)
}
$MailMessage.Subject = "Inventario SOLEX - $($objWMIService.Name) - Service Tag: $serviceTag - Usuario: $userSession"
$MailMessage.Body = "Adjunto encontrarÃ¡s el inventario del equipo."
$MailMessage.IsBodyHtml = $false
$MailMessage.Attachments.Add((New-Object System.Net.Mail.Attachment($strFilePath)))

$SmtpClient = New-Object System.Net.Mail.SmtpClient($smtp_server, $smtp_port)
$SmtpClient.Credentials = $cred
$SmtpClient.EnableSsl = $true
$SmtpClient.Send($MailMessage)


Write-Host "Correo enviado exitosamente."

