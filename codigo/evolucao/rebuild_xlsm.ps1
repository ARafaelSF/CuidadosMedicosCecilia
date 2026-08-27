$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$codigo = Split-Path -Parent $here
$root = Split-Path -Parent $codigo
$pathXlsx = Join-Path $env:TEMP "EvolucaoExames.xlsx"
$pathXlsmTmp = Join-Path $env:TEMP "EvolucaoExames.xlsm"
$pathXlsm = [System.IO.Path]::Combine($root, "Exames", "Evolução Exames.xlsm")
$modBas = Join-Path $here "Marcacao.bas"
$wbBas = Join-Path $here "ThisWorkbook.bas"

if (-not (Test-Path -LiteralPath $pathXlsx)) {
    throw "Falta $pathXlsx. Rode antes: py -3 build_evolucao.py"
}

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
    $wb = $excel.Workbooks.Open($pathXlsx)
    $ws = $wb.Worksheets.Item("Escolher")

    try {
        while ($ws.CheckBoxes().Count -gt 0) { $ws.CheckBoxes().Item(1).Delete() }
    } catch {}

    $first = 6
    $last = $ws.Cells.Item($ws.Rows.Count, 4).End(-4162).Row
    $xlOff = -4146
    for ($row = $first; $row -le $last; $row++) {
        $titulo = [string]$ws.Cells.Item($row, 4).Value2
        if ([string]::IsNullOrWhiteSpace($titulo)) { continue }
        $left = [double]$ws.Cells.Item($row, 1).Left + 5
        $top = [double]$ws.Cells.Item($row, 1).Top + 3
        $cb = $ws.CheckBoxes().Add($left, $top, 18, 18)
        $cb.Caption = ""
        $cb.LinkedCell = "B$row"
        $cb.Value = $xlOff
        $ws.Cells.Item($row, 2).Value2 = $false
        $cb.Display3DShading = $false
        try { $cb.PrintObject = $false } catch {}
        # Sem OnAction: clique so marca a celula; filtragem ao abrir abas Selecionados
        try { $cb.OnAction = "NopSelecao" } catch {}
    }
    Write-Host "boxes $($ws.CheckBoxes().Count)"

    try {
        while ($ws.Buttons().Count -gt 0) { $ws.Buttons().Item(1).Delete() }
    } catch {}
    try {
        while ($ws.OptionButtons().Count -gt 0) { $ws.OptionButtons().Item(1).Delete() }
    } catch {}

    $rowCtrl = 3
    $top = [double]$ws.Cells.Item($rowCtrl, 6).Top
    $leftF = [double]$ws.Cells.Item($rowCtrl, 6).Left
    $btn1 = $ws.Buttons().Add($leftF, $top, 150, 22)
    $btn1.Caption = "Imprimir selecionados"
    $btn1.OnAction = "ImprimirGraficosSelecionados"
    $btn2 = $ws.Buttons().Add($leftF, $top + 26, 150, 22)
    $btn2.Caption = "Imprimir todos"
    $btn2.OnAction = "ImprimirGraficosTodos"

    $xlOn = 1
    $leftOb = $leftF + 165
    $ob3 = $ws.OptionButtons().Add($leftOb, $top, 115, 18)
    $ob3.Caption = "3 por folha"
    $ob3.LinkedCell = "`$H`$3"
    $ob3.Display3DShading = $false
    $ob4 = $ws.OptionButtons().Add($leftOb, $top + 22, 115, 18)
    $ob4.Caption = "4 por folha"
    $ob4.LinkedCell = "`$H`$3"
    $ob4.Display3DShading = $false
    $ob3.Value = $xlOn
    $ws.Cells.Item($rowCtrl, 8).Value2 = 1
    Write-Host "print_buttons"

    $vbaOk = $false
    try {
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
        $tw = $proj.VBComponents.Item("EstaPastaDeTrabalho").CodeModule
        if ($tw.CountOfLines -gt 0) { $tw.DeleteLines(1, $tw.CountOfLines) }
        $tw.AddFromFile($wbBas)
        $vbaOk = $true
        Write-Host "VBA_OK"
    } catch {
        Write-Host "VBA_FAIL $($_.Exception.Message)"
    }

    $count = $ws.CheckBoxes().Count
    for ($i = 1; $i -le $count; $i++) {
        try { $ws.CheckBoxes().Item($i).OnAction = "NopSelecao" } catch {}
    }

    if ($vbaOk) {
        try {
            $excel.Run("LimparOnActionCaixas")
            $excel.Run("AtualizarImpressao")
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
        $copied = $false
        for ($t = 1; $t -le 5; $t++) {
            try {
                Copy-Item -LiteralPath $pathXlsmTmp -Destination $pathXlsm -Force
                $copied = $true
                break
            } catch {
                Start-Sleep -Seconds 2
            }
        }
        if (-not $copied) {
            $alt = [System.IO.Path]::Combine($root, "Exames", "Evolucao Exames novo.xlsm")
            Copy-Item -LiteralPath $pathXlsmTmp -Destination $alt -Force
            Write-Host "SAVED_ALT $alt"
        } else {
            Write-Host "SAVED_XLSM"
        }
    } else {
        $wb.Save()
        Write-Host "SAVED_XLSX"
        $wb.Close($true)
    }
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
