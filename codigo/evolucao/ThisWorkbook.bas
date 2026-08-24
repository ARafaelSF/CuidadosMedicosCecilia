Private Sub Workbook_Open()
    On Error Resume Next
    Dim ws As Worksheet
    Dim cb As CheckBox
    Dim btn As Button
    Set ws = Me.Worksheets("Escolher")
    For Each cb In ws.CheckBoxes
        cb.OnAction = "AtualizarImpressao"
    Next cb
    For Each btn In ws.Buttons
        If InStr(1, btn.Caption, "todos", vbTextCompare) > 0 Then
            btn.OnAction = "ImprimirGraficosTodos"
        ElseIf InStr(1, btn.Caption, "Imprimir", vbTextCompare) > 0 Then
            btn.OnAction = "ImprimirGraficosSelecionados"
        End If
    Next btn
    AtualizarImpressao
    EsconderSeriesInvisiveis
    AjustarAlturasDados
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If Sh.Name <> "Escolher" Then Exit Sub
    If Intersect(Target, Sh.Columns(2)) Is Nothing Then Exit Sub
    AtualizarImpressao
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    If Sh.Name = "Dados Completo" Or Sh.Name = "Dados Selecionados" Then
        AjustarAlturasDados
    End If
End Sub
