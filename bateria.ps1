# Comando para obtener el n�mero de serie del equipo con PowerShell
$powershellCommand = "(Get-WmiObject -Class Win32_Bios).SerialNumber"
$serialNumber = Invoke-Expression -Command $powershellCommand

# Comando a ejecutar en PowerShell para generar el informe de bater�a
$batteryReportCommand = "powercfg /batteryreport"
Invoke-Expression -Command $batteryReportCommand

# Obtener la ruta de %userprofile%
$userProfilePath = [System.Environment]::GetFolderPath('UserProfile')

# Mover el archivo del informe de bater�a a la ruta de %userprofile%
$sourceFile = Join-Path -Path $userProfilePath -ChildPath "battery-report.html"
$destinationFile = Join-Path -Path $userProfilePath -ChildPath "battery-report.html"

if (Test-Path $sourceFile) {
    Move-Item -Path $sourceFile -Destination $destinationFile
}
else {
    Write-Host "El archivo del informe de bater�a no se encontr� en la ubicaci�n especificada: $sourceFile"
}

# Configurar detalles del correo electr�nico
$smtpServer = "smtp.gmail.com"
$smtpPort = 587
$smtpUsername = "solexactivosscl@gmail.com"
$smtpPassword = ConvertTo-SecureString "fznm abuv tsjl nbcj" -AsPlainText -Force
$emailFrom = "solexactivosscl@gmail.com"
$emailTo = "tickets@solex.biz"
$emailSubject = "Inventario SOLEX"
$emailBody = "Adjunto encontrar�s el inventario del equipo."

# Enviar el correo electr�nico con el archivo adjunto
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
