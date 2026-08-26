Private Sub Workbook_Open()
    On Error Resume Next
    AjustarAlturasResumo
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    If Sh.Name = "Resumo" Then AjustarAlturasResumo
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Static busy As Boolean
    If busy Then Exit Sub
    If Sh.Name <> "Resumo" Then Exit Sub
    busy = True
    On Error GoTo Limpa
    AjustarAlturasResumo
Limpa:
    busy = False
End Sub
