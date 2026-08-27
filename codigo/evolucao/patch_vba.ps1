# Patch só o VBA + OnAction do Evolução Exames.xlsm (sem rebuild completo).
# Gera caminhos UTF-8 via Python (evita problema de acentos no PowerShell).
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$pathsFile = Join-Path $env:TEMP "evol_patch_paths.txt"
py -3 -c @"
from pathlib import Path
import sys
sys.path.insert(0, r'$($here | Split-Path -Parent)')
from caminhos import DEST_EVOL_XLSM, CODIGO
Path(r'$pathsFile').write_text(
    str(DEST_EVOL_XLSM) + chr(10) +
    str(CODIGO / 'evolucao' / 'Marcacao.bas') + chr(10) +
    str(CODIGO / 'evolucao' / 'ThisWorkbook.bas') + chr(10),
    encoding='utf-8')
print(DEST_EVOL_XLSM)
"@

$paths = Get-Content -LiteralPath $pathsFile -Encoding UTF8
$pathXlsm = $paths[0].Trim()
$modBas = $paths[1].Trim()
$wbBas = $paths[2].Trim()

if (-not (Test-Path -LiteralPath $pathXlsm)) { throw "Falta $pathXlsm" }
Write-Host "xlsm $pathXlsm"

Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 1

$reg = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
$oldVbom = $null
if (Test-Path $reg) {
    $oldVbom = (Get-ItemProperty -Path $reg -ErrorAction SilentlyContinue).AccessVBOM
    New-ItemProperty -Path $reg -Name AccessVBOM -Value 1 -PropertyType DWord -Force | Out-Null
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false

try {
    $wb = $excel.Workbooks.Open($pathXlsm)
    $ws = $wb.Worksheets.Item("Escolher")

    $proj = $wb.VBProject
    $hasMarc = $false
    foreach ($c in @($proj.VBComponents)) {
        if ($c.Name -eq "Marcacao") { $hasMarc = $true }
    }
    if (-not $hasMarc) {
        $std = $proj.VBComponents.Add(1)
        $std.Name = "Marcacao"
    }
    $m = $proj.VBComponents.Item("Marcacao").CodeModule
    if ($m.CountOfLines -gt 0) { $m.DeleteLines(1, $m.CountOfLines) }
    $m.AddFromFile($modBas)

    $twName = $null
    foreach ($c in @($proj.VBComponents)) {
        if ($c.Type -eq 100) { $twName = $c.Name; break }
    }
    if (-not $twName) { $twName = "EstaPastaDeTrabalho" }
    $tw = $proj.VBComponents.Item($twName).CodeModule
    if ($tw.CountOfLines -gt 0) { $tw.DeleteLines(1, $tw.CountOfLines) }
    $tw.AddFromFile($wbBas)
    Write-Host "VBA_OK ($twName)"

    $count = $ws.CheckBoxes().Count
    for ($i = 1; $i -le $count; $i++) {
        try { $ws.CheckBoxes().Item($i).OnAction = "" } catch {}
    }
    Write-Host "cleared OnAction on $count boxes"

    try {
        $excel.Run("LimparOnActionCaixas")
        $excel.Run("AtualizarImpressao")
        Write-Host "MACRO_RUN"
    } catch {
        Write-Host "MACRO_FAIL $($_.Exception.Message)"
    }

    $wb.Save()
    $wb.Close($true)
    Write-Host "SAVED"
}
finally {
    $excel.ScreenUpdating = $true
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force
    if (Test-Path $reg) {
        if ($null -eq $oldVbom) {
            Remove-ItemProperty -Path $reg -Name AccessVBOM -ErrorAction SilentlyContinue
        } else {
            Set-ItemProperty -Path $reg -Name AccessVBOM -Value $oldVbom
        }
    }
}
