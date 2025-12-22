# ==========================================
# Gestion DCE/RPC - Firewall Windows
# ==========================================

$ruleName = "Block DCE-RPC TCP 135"

# ---- Verificar permisos de Administrador ----
function Test-Admin {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Admin)) {
    Write-Host "ERROR: Este script debe ejecutarse como ADMINISTRADOR." -ForegroundColor Red
    Pause
    Exit
}

# ---- Menu ----
function Show-Menu {
    Clear-Host
    Write-Host "=====================================" -ForegroundColor Cyan
    Write-Host "  Gestion DCE/RPC - Firewall Windows  " -ForegroundColor Cyan
    Write-Host "=====================================" -ForegroundColor Cyan
    Write-Host "1. Bloquear DCE/RPC (TCP 135)"
    Write-Host "2. Desbloquear DCE/RPC (Eliminar regla)"
    Write-Host "3. Verificar estado de la regla"
    Write-Host "4. Salir"
    Write-Host ""
}

# ---- Bloquear ----
function Block-DCE {
    try {
        if (Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue) {
            Write-Host "La regla ya existe." -ForegroundColor Yellow
        } else {
            New-NetFirewallRule `
                -DisplayName $ruleName `
                -Direction Inbound `
                -Protocol TCP `
                -LocalPort 135 `
                -Action Block `
                -Profile Any `
                -ErrorAction Stop

            Write-Host "DCE/RPC TCP 135 BLOQUEADO correctamente." -ForegroundColor Green
        }
    } catch {
        Write-Host "ERROR al crear la regla:" -ForegroundColor Red
        Write-Host $_.Exception.Message
    }
    Pause
}

# ---- Desbloquear ----
function Unblock-DCE {
    try {
        $rule = Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue
        if ($rule) {
            Remove-NetFirewallRule -DisplayName $ruleName -ErrorAction Stop
            Write-Host "DCE/RPC TCP 135 DESBLOQUEADO (regla eliminada)." -ForegroundColor Green
        } else {
            Write-Host "No existe una regla para eliminar." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "ERROR al eliminar la regla:" -ForegroundColor Red
        Write-Host $_.Exception.Message
    }
    Pause
}

# ---- Verificar Estado ----
function Check-DCEStatus {
    Clear-Host
    Write-Host "Estado actual de la regla DCE/RPC" -ForegroundColor Cyan
    Write-Host "---------------------------------" -ForegroundColor Cyan

    $rule = Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue
    if ($rule) {
        $rule | Select-Object DisplayName, Enabled, Direction, Action, Profile
    } else {
        Write-Host "La regla NO existe." -ForegroundColor Yellow
    }
    Write-Host ""
    Pause
}

# ---- Loop principal ----
do {
    Show-Menu
    $option = Read-Host "Seleccione una opcion"

    switch ($option) {
        "1" { Block-DCE }
        "2" { Unblock-DCE }
        "3" { Check-DCEStatus }
        "4" { Write-Host "Saliendo..." -ForegroundColor Cyan }
        default {
            Write-Host "Opcion invalida." -ForegroundColor Red
            Pause
        }
    }
} while ($option -ne "4")
