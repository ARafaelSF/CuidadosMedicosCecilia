$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$codigo = Split-Path -Parent $here
$root = Split-Path -Parent $codigo
$pathXlsx = [System.IO.Path]::Combine($root, "Exames", "Resumo Exames.xlsx")
$pathXlsm = [System.IO.Path]::Combine($root, "Exames", "Resumo Exames.xlsm")
$pathXlsmTmp = Join-Path $env:TEMP "ResumoExames.xlsm"
$modBas = Join-Path $here "Alturas.bas"
$wbBas = Join-Path $here "EstaPastaDeTrabalho.bas"

# Preferir xlsx se existir; senão reabrir o xlsm para reinstalar VBA
$src = $null
if (Test-Path -LiteralPath $pathXlsx) { $src = $pathXlsx }
elseif (Test-Path -LiteralPath $pathXlsm) { $src = $pathXlsm }
else { throw "Não achei Resumo Exames.xlsx nem .xlsm em Exames\" }

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
    $wb = $excel.Workbooks.Open($src)
    $ws = $wb.Worksheets.Item("Resumo")

    # Garante quebra de texto na Descrição (coluna D = 4)
    $last = $ws.Cells.Item($ws.Rows.Count, 1).End(-4162).Row
    if ($last -ge 3) {
        $ws.Range("D3:D$last").WrapText = $true
        $ws.Range("D3:D$last").VerticalAlignment = -4108  # xlCenter
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
        $wb.SaveAs($pathXlsmTmp, 52)  # xlOpenXMLWorkbookMacroEnabled
        Write-Host "SAVED_TMP"
        $wb.Close($true)
        Start-Sleep -Seconds 1
        Copy-Item -LiteralPath $pathXlsmTmp -Destination $pathXlsm -Force
        Write-Host "SAVED_XLSM $pathXlsm"
        # Evita dois arquivos: o ativo passa a ser o .xlsm
        if ((Test-Path -LiteralPath $pathXlsx) -and ($src -eq $pathXlsx)) {
            Remove-Item -LiteralPath $pathXlsx -Force
            Write-Host "REMOVED_XLSX"
        }
    } else {
        $wb.Close($false)
        throw "Não consegui instalar o VBA (marque 'Confiar no acesso ao modelo de objeto do projeto do VBA' nas opções do Excel)."
    }
}
finally {
    $excel.Quit() | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    if ($null -ne $oldVbom) {
        Set-ItemProperty -Path $reg -Name AccessVBOM -Value $oldVbom
    } elseif (Test-Path $reg) {
        Remove-ItemProperty -Path $reg -Name AccessVBOM -ErrorAction SilentlyContinue
    }
}

Write-Host "Pronto. Abra Resumo Exames.xlsm em Exames\ (habilitar macros)."
