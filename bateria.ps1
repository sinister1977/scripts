# Comando para obtener el número de serie del equipo con PowerShell
$powershellCommand = "(Get-WmiObject -Class Win32_Bios).SerialNumber"
$serialNumber = Invoke-Expression -Command $powershellCommand

# Comando a ejecutar en PowerShell para generar el informe de batería
$batteryReportCommand = "powercfg /batteryreport"
Invoke-Expression -Command $batteryReportCommand

# Obtener la ruta de %userprofile%
$userProfilePath = [System.Environment]::GetFolderPath('UserProfile')

# Mover el archivo del informe de batería a la ruta de %userprofile%
$sourceFile = Join-Path -Path $userProfilePath -ChildPath "battery-report.html"
$destinationFile = Join-Path -Path $userProfilePath -ChildPath "battery-report.html"

if (Test-Path $sourceFile) {
    Move-Item -Path $sourceFile -Destination $destinationFile
}
else {
    Write-Host "El archivo del informe de batería no se encontró en la ubicación especificada: $sourceFile"
}

# Configurar detalles del correo electrónico
$smtpServer = (-join [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('c210cC1yZWxheS5icmV2by5jb20=')).ToCharArray() | ForEach-Object {$_}) -join ''
$smtpPort = 587
$smtpUsername = (-join [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('c2EyMjRhYTAwMUBzbXRwLWJyZXZvLmNvbQ==')).ToCharArray() | ForEach-Object {$_}) -join ''
$smtpPassword = (-join [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('eHNtdHBzaWItOGIyYTRmODA2ZDQ0NzJiYmQ4M2RhYjIzNTFiYmMyNTIxN2I2OWYxNmNiYmMzMTRmYjZlZjViNzY0YWZhYjBhMC1scEgyTXJlU1hlaEw1OVRY')).ToCharArray() | ForEach-Object {$_}) -join ''
$emailFrom = (-join [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('c29sZXhhY3Rpdm9zc2NsQGdtYWlsLmNvbQ==')).ToCharArray() | ForEach-Object {$_}) -join ''
$emailTo = (-join [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('c2VyZ2lvLnNhbm1hcnRpbkBzb2xleC5iaXo=')).ToCharArray() | ForEach-Object {$_}) -join ''
$emailSubject = (-join [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('SW52ZW50YXJpbyBCYXRlcmlhIFNPTEVY')).ToCharArray() | ForEach-Object {$_}) -join ''
$emailBody = (-join [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('QWRqdW50byBlbmNvbnRyYXLDqXMgZWwgcmVwb3J0ZSBkZSBiYXRlcmlhIFNPTEVYIC4=')).ToCharArray() | ForEach-Object {$_}) -join ''

# Enviar el correo electrónico con el archivo adjunto
$mailParams = @{
    SmtpServer        = $smtpServer
    Port              = $smtpPort
    UseSsl            = $true
    Credential        = New-Object System.Management.Automation.PSCredential($smtpUsername, $smtpPassword)
    From              = $emailFrom
    To                = $emailTo
    Subject           = $emailSubject
    Body              = $emailBody
    Attachments       = $destinationFile
}

Send-MailMessage @mailParams

