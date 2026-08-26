$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$pathXlsxTmp = Join-Path $env:TEMP "ResumoRelatorios.xlsx"
$pathXlsmTmp = Join-Path $env:TEMP "ResumoRelatorios.xlsm"
$modBas = Join-Path $here "AlturasRelatorios.bas"
$wbBas = Join-Path $here "EstaPastaDeTrabalhoRelatorios.bas"

# Destinos Unicode via caminhos.py (evita mojibake do PowerShell)
$destLines = py -3 -c @"
import sys
from pathlib import Path
sys.path.insert(0, str(Path(r'$here').parent))
from caminhos import RELATORIOS
print(RELATORIOS / 'Resumo Relatórios.xlsx')
print(RELATORIOS / 'Resumo Relatórios.xlsm')
print(RELATORIOS)
"@
$pathXlsx = $destLines[0]
$pathXlsm = $destLines[1]

$src = $null
if (Test-Path -LiteralPath $pathXlsxTmp) { $src = $pathXlsxTmp }
elseif (Test-Path -LiteralPath $pathXlsx) { $src = $pathXlsx }
elseif (Test-Path -LiteralPath $pathXlsm) { $src = $pathXlsm }
else { throw "Não achei ResumoRelatorios.xlsx / Resumo Relatórios.xlsx / .xlsm" }
Write-Host "SRC $src"
Write-Host "DEST_XLSM_TMP $pathXlsmTmp"

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
    if (-not (Test-Path -LiteralPath $src)) { throw "Arquivo fonte inexistente: $src" }
    Write-Host "Opening..."
    $wb = $excel.Workbooks.Open($src, 0, $false)
    if ($null -eq $wb) { throw "Workbooks.Open retornou nulo para $src" }
    $ws = $wb.Worksheets.Item("Resumo")

    $last = 2
    for ($r = 3; $r -le 400; $r++) {
        $tipo = [string]$ws.Cells.Item($r, 2).Value2
        if ([string]::IsNullOrWhiteSpace($tipo)) { break }
        $last = $r
    }
    if ($last -ge 3) {
        $ws.Range("D3:D$last").WrapText = $true
        $ws.Range("D3:D$last").VerticalAlignment = -4108
        try {
            while ($ws.ListObjects.Count -gt 0) { $ws.ListObjects.Item(1).Delete() }
        } catch {}
        $rng = $ws.Range("A2:G$last")
        $lo = $ws.ListObjects.Add(1, $rng, $null, 1)
        $lo.Name = "TabelaRelatorios"
        $xlNone = -4142
        for ($r = 3; $r -le $last; $r++) {
            foreach ($c in @(1, 3, 4, 5, 6, 7)) {
                $ws.Cells.Item($r, $c).Interior.Pattern = $xlNone
            }
        }
        $lo.TableStyle = "TableStyleMedium2"
        $lo.ShowTableStyleRowStripes = $true
        $lo.ShowTableStyleColumnStripes = $false
        $ws.Columns.Item(5).ColumnWidth = 60
        Write-Host "TABLE_OK rows=$($lo.ListRows.Count) last=$last stripes=$($lo.ShowTableStyleRowStripes)"
    }

    $vbaOk = $false
    try {
        $proj = $wb.VBProject
        $has = $false
        foreach ($c in @($proj.VBComponents)) {
            if ($c.Name -eq "Alturas") { $has = $true }
        }
        if (-not $has) {
            $std = $proj.VBComponents.Add(1)
            $std.Name = "Alturas"
        }
        $m = $proj.VBComponents.Item("Alturas").CodeModule
        if ($m.CountOfLines -gt 0) { $m.DeleteLines(1, $m.CountOfLines) }
        $m.AddFromFile($modBas)

        $tw = $proj.VBComponents.Item("EstaPastaDeTrabalho").CodeModule
        if ($tw.CountOfLines -gt 0) { $tw.DeleteLines(1, $tw.CountOfLines) }
        $tw.AddFromFile($wbBas)
        $vbaOk = $true
        Write-Host "VBA_OK"
    } catch {
        Write-Host "VBA_FAIL $($_.Exception.Message)"
    }

    if ($vbaOk) {
        try {
            $excel.Run("AjustarAlturasResumo")
            Write-Host "MACRO_RUN"
        } catch {
            Write-Host "MACRO_FAIL $($_.Exception.Message)"
        }
    }

    if ($vbaOk) {
        if (Test-Path -LiteralPath $pathXlsmTmp) { Remove-Item -LiteralPath $pathXlsmTmp -Force }
        $wb.SaveAs($pathXlsmTmp, 52)
        Write-Host "SAVED_TMP"
        $wb.Close($true)
        Start-Sleep -Seconds 1
        py -3 (Join-Path $here "copy_resumo_relatorios.py")

    } else {
        $wb.Close($false)
        throw "Não consegui instalar o VBA."
    }
}
finally {
    try { $excel.Quit() | Out-Null } catch {}
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    if ($null -ne $oldVbom) {
        Set-ItemProperty -Path $reg -Name AccessVBOM -Value $oldVbom
    } elseif (Test-Path $reg) {
        Remove-ItemProperty -Path $reg -Name AccessVBOM -ErrorAction SilentlyContinue
    }
}

Write-Host "Pronto. Abra Resumo Relatórios.xlsm em Relatórios\ (habilitar macros)."
