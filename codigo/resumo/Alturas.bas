Option Explicit

' Ajusta a altura de cada linha da TabelaExames ao texto da Descrição.
' Chamar depois de ordenar/filtrar (SelectionChange) e ao abrir o arquivo.

Public Sub AjustarAlturasResumo()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim i As Long
    Dim rng As Range
    Dim h As Double

    On Error GoTo Fim
    Set ws = ThisWorkbook.Worksheets("Resumo")
    Set lo = ws.ListObjects("TabelaExames")
    If lo Is Nothing Then GoTo Fim
    If lo.DataBodyRange Is Nothing Then GoTo Fim

    Application.ScreenUpdating = False
    For i = 1 To lo.ListRows.Count
        Set rng = lo.ListRows(i).Range.EntireRow
        ' Reseta e recalcula (AutoFit respeita quebra de texto na coluna Descrição)
        rng.AutoFit
        h = rng.RowHeight
        If h < 18 Then rng.RowHeight = 18
        If h > 140 Then rng.RowHeight = 140
    Next i

Fim:
    Application.ScreenUpdating = True
End Sub
