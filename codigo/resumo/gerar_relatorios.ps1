$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 1
Write-Host "1/2 rebuild xlsx..."
py -3 (Join-Path $here "rebuild_relatorios.py")
Write-Host "2/2 VBA xlsm..."
powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $here "rebuild_relatorios_xlsm.ps1")
Write-Host "Pronto. Abra Resumo Relatórios.xlsm em Relatórios\ (habilitar macros)."
