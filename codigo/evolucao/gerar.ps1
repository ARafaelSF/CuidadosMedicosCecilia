$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 1
Write-Host "1/2 gerando xlsx..."
py -3 (Join-Path $here "build_evolucao.py")
Write-Host "2/2 caixinhas + VBA..."
powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $here "rebuild_xlsm.ps1")
Write-Host "Pronto. Abra Evolução Exames.xlsm em Exames\ (habilitar macros)."
