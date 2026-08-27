Private Sub Workbook_Open()
    On Error Resume Next
    Dim ws As Worksheet
    Dim btn As Button
    Application.EnableEvents = True
    Set ws = Me.Worksheets("Escolher")
    LimparOnActionCaixas
    For Each btn In ws.Buttons
        If InStr(1, btn.Caption, "todos", vbTextCompare) > 0 Then
            btn.OnAction = "ImprimirGraficosTodos"
        ElseIf InStr(1, btn.Caption, "Imprimir", vbTextCompare) > 0 Then
            btn.OnAction = "ImprimirGraficosSelecionados"
        End If
    Next btn
    ' Nao roda AtualizarImpressao no open (pesado). Filtra ao abrir abas Selecionados.
    EsconderSeriesInvisiveis
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    ' Intencionalmente vazio: LinkedCell das caixas nao deve filtrar a cada clique.
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    If Sh.Name = "Dados Selecionados" Or Sh.Name = "Graficos Selecionados" Then
        Application.StatusBar = "Atualizando selecao..."
        AtualizarImpressao
        Application.StatusBar = False
    End If
End Sub
