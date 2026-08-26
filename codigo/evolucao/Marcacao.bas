Option Explicit

Private Const msoTrueVal As Long = -1
Private Const msoFalseVal As Long = 0
Private Const xlMoveAndSize As Long = 1
Private Const xlVeryHidden As Long = 2
Private Const xlPortrait As Long = 1
Private Const xlPaperA4 As Long = 9
Private Const xlTypePDF As Long = 0
Private Const xlQualityStandard As Long = 0
Private Const xlScreen As Long = 1
Private Const xlPictureFmt As Long = -4147
Private Const xlCalculationManual As Long = -4135
Private Const xlColorIndexNone As Long = -4142

Private Const ABA_ESCOLHER As String = "Escolher"
Private Const ABA_DADOS_COMP As String = "Dados Completo"
Private Const ABA_DADOS_SEL As String = "Dados Selecionados"
Private Const ABA_GRAF_COMP As String = "Graficos Completo"
Private Const ABA_GRAF_SEL As String = "Graficos Selecionados"
Private Const ABA_MAPA As String = "Mapa"
Private Const ABA_UNICOS As String = "Unicos"

Private Function Folha(nome As String) As Worksheet
    On Error Resume Next
    Set Folha = ThisWorkbook.Worksheets(nome)
    On Error GoTo 0
End Function

Public Sub AtualizarImpressao()
    Dim wsS As Worksheet, wsC As Worksheet, wsG As Worksheet
    Dim wsM As Worksheet, wsE As Worksheet
    Dim last As Long, i As Long, er As Long
    Dim r1 As Long, r2 As Long, g1 As Long, g2 As Long, r As Long
    Dim v As Variant, show As Boolean

    On Error GoTo Falha
    Set wsS = Folha(ABA_DADOS_SEL)
    Set wsC = Folha(ABA_DADOS_COMP)
    Set wsM = Folha(ABA_MAPA)
    Set wsE = Folha(ABA_ESCOLHER)
    Set wsG = Folha(ABA_GRAF_SEL)
    If wsS Is Nothing Or wsM Is Nothing Or wsE Is Nothing Then
        Err.Raise 9, , "Aba nao encontrada. Nomes esperados: " & ABA_DADOS_SEL & " / " & ABA_ESCOLHER
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    wsS.Rows.Hidden = False
    If Not wsG Is Nothing Then wsG.Rows.Hidden = False
    If Not wsG Is Nothing Then GarantirNomesGraficos wsG

    last = wsM.Cells(wsM.Rows.Count, 1).End(xlUp).Row
    For i = 2 To last
        er = CLng(wsM.Cells(i, 1).Value)
        r1 = CLng(wsM.Cells(i, 2).Value)
        r2 = CLng(wsM.Cells(i, 3).Value)
        g1 = 0: g2 = 0
        If Not IsEmpty(wsM.Cells(i, 4).Value) Then g1 = CLng(wsM.Cells(i, 4).Value)
        If Not IsEmpty(wsM.Cells(i, 5).Value) Then g2 = CLng(wsM.Cells(i, 5).Value)
        v = wsE.Cells(er, 2).Value
        show = CelulaMarcada(v)
        If r1 > 0 And r2 >= r1 Then
            For r = r1 To r2
                wsS.Rows(r).Hidden = Not show
            Next r
        End If
        If Not wsG Is Nothing Then
            If g1 > 0 And g2 >= g1 Then
                For r = g1 To g2
                    wsG.Rows(r).Hidden = Not show
                Next r
            End If
        End If
    Next i

    If Not wsG Is Nothing Then AplicarGraficos wsG, wsM, wsE, last

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
Falha:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Sub ImprimirGraficosSelecionados()
    ImprimirGraficos False
End Sub

Public Sub ImprimirGraficosTodos()
    ImprimirGraficos True
End Sub

Public Sub ImprimirGraficosMarcados()
    ImprimirGraficos False
End Sub

Private Function PorFolha() As Long
    Dim v As Variant
    Dim wsE As Worksheet
    On Error Resume Next
    Set wsE = Folha(ABA_ESCOLHER)
    If Not wsE Is Nothing Then v = wsE.Range("H3").Value
    On Error GoTo 0
    If IsNumeric(v) Then
        If CLng(v) = 2 Or CLng(v) = 4 Then
            PorFolha = 4
            Exit Function
        End If
    End If
    PorFolha = 3
End Function

Private Sub ImprimirGraficos(todos As Boolean)
    Dim wsG As Worksheet, wsP As Worksheet, wsE As Worksheet
    Dim wsM As Worksheet, ws As Worksheet, wsOld As Worksheet
    Dim wsFull As Worksheet
    Dim last As Long, lastP As Long, i As Long, n As Long, r As Long
    Dim origRow As Long, nm As String, titulo As String
    Dim co As ChartObject
    Dim destTmp As String
    Dim porPagina As Long
    Dim picW As Double, picH As Double, titleH As Double
    Dim wsPObj As Worksheet
    Dim idx As Long
    Dim inclui As Boolean
    Dim errN As Long
    Dim errD As String
    Dim pdfOk As Boolean
    Dim fonteImp As Long
    Dim nUni As Long
    Dim calcAnt As Long
    Dim usableW As Double

    porPagina = PorFolha()

    errN = 0
    calcAnt = Application.Calculation
    Set wsOld = ActiveSheet
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo FalhaImp

    AtualizarImpressao

    Set wsE = Folha(ABA_ESCOLHER)
    Set wsM = Folha(ABA_MAPA)
    Set wsG = Folha(ABA_GRAF_SEL)
    Set wsFull = Folha(ABA_GRAF_COMP)
    If wsG Is Nothing Then
        On Error Resume Next
        Application.Calculation = calcAnt
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Aviso "Aba '" & ABA_GRAF_SEL & "' nao encontrada."
        Exit Sub
    End If
    If wsFull Is Nothing Then Set wsFull = wsG

    calcAnt = Application.Calculation
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    Err.Clear
    On Error GoTo FalhaImp

    last = wsM.Cells(wsM.Rows.Count, 1).End(xlUp).Row
    Set wsPObj = FolhaPosicoes()
    lastP = wsPObj.Cells(wsPObj.Rows.Count, 1).End(xlUp).Row

    Set wsP = FolhaImpressao(porPagina)
    wsP.Activate
    ' Layout estável: 3 ou 4 gráficos por folha A4 (título + gráfico)
    titleH = 16
    ' Margens reais (após PageSetup aplicado) — define a largura de UMA página
    usableW = Application.CentimetersToPoints(21) - wsP.PageSetup.LeftMargin - wsP.PageSetup.RightMargin
    If usableW < 400 Then usableW = 400
    ' Colunas A–F cabem numa página; gráfico = largura dessas colunas
    ConfigurarColunasImpressao wsP, usableW
    picW = SomaLargurasColunas(wsP, 1, 6) - 8
    If picW < 280 Then picW = 280
    picH = (Application.CentimetersToPoints(29.7) - wsP.PageSetup.TopMargin - wsP.PageSetup.BottomMargin) / porPagina - titleH - 6
    If picH < 130 Then picH = 130
    fonteImp = 9
    If porPagina = 4 Then fonteImp = 8

    r = 1
    n = 0
    For i = 2 To lastP
        origRow = CLng(wsPObj.Cells(i, 2).Value)
        nm = CStr(wsPObj.Cells(i, 1).Value)
        If origRow <= 0 Or Len(nm) = 0 Then GoTo ProxImp
        If todos Then
            inclui = True
        Else
            inclui = GraficoMarcado(wsM, wsE, last, origRow)
        End If
        If Not inclui Then GoTo ProxImp

        Set co = GraficoNaLinha(wsFull, origRow)
        If co Is Nothing Then Set co = GraficoNaLinha(wsG, origRow)
        If co Is Nothing Then
            On Error Resume Next
            Set co = wsFull.ChartObjects(nm)
            If co Is Nothing Then Set co = wsG.ChartObjects(nm)
            Err.Clear
        End If
        On Error GoTo FalhaImp
        If co Is Nothing Then GoTo ProxImp

        If n > 0 And (n Mod porPagina) = 0 Then
            On Error Resume Next
            wsP.Rows(r).PageBreak = -4135
            Err.Clear
            On Error GoTo FalhaImp
        End If

        titulo = Trim$(CStr(wsFull.Cells(origRow - 2, 1).Value))
        If Len(titulo) = 0 Then titulo = Trim$(CStr(wsG.Cells(origRow - 2, 1).Value))
        If Len(titulo) = 0 Then titulo = Trim$(CStr(wsG.Cells(origRow - 1, 1).Value))
        If Len(titulo) = 0 Then titulo = nm
        wsP.Rows(r).RowHeight = titleH
        wsP.Rows(r + 1).RowHeight = picH

        ' Título mesclado A:F (mesma largura da tabela de únicos)
        On Error Resume Next
        wsP.Range(wsP.Cells(r, 1), wsP.Cells(r, 6)).UnMerge
        Err.Clear
        On Error GoTo FalhaImp
        wsP.Range(wsP.Cells(r, 1), wsP.Cells(r, 6)).Merge
        With wsP.Cells(r, 1)
            .Value = titulo
            .Font.Bold = True
            .Font.Size = 10
            .Font.Name = "Calibri"
            .Font.Color = RGB(31, 78, 121)
            .Interior.Color = RGB(217, 234, 247)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With

        If Not ColocarGraficoNaFolha(co, wsP, 2, wsP.Cells(r + 1, 1).Top + 1, picW, picH - 4, fonteImp) Then
            wsP.Rows(r).RowHeight = 15
            wsP.Rows(r + 1).RowHeight = 15
            wsP.Cells(r, 1).Value = titulo & " (falha ao colar grafico)"
            GoTo ProxImp
        End If

        n = n + 1
        r = r + 2
        Application.StatusBar = "Montando impressao " & CStr(n) & "..."
ProxImp:
    Next i
    Application.StatusBar = False

    nUni = ColarTabelaUnicos(wsP, wsE, todos, n, r, porPagina)

    If n = 0 And nUni = 0 Then
        On Error Resume Next
        Application.Calculation = calcAnt
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Application.StatusBar = False
        If Not wsOld Is Nothing Then wsOld.Activate
        If todos Then
            Aviso "Nenhum grafico para imprimir."
        Else
            Aviso "Marque ao menos um exame com grafico, ou use Imprimir todos."
        End If
        Exit Sub
    End If

    On Error Resume Next
    ' Garante que nada ultrapasse a largura de uma página (sem quebra vertical)
    CaberNaLarguraDaPagina wsP, usableW
    wsP.PageSetup.PrintArea = "$A$1:$F$" & CStr(Application.Max(r - 1, 1))
    wsP.PageSetup.Zoom = False
    wsP.PageSetup.FitToPagesWide = 1
    wsP.PageSetup.FitToPagesTall = False
    Application.CutCopyMode = False
    wsP.Activate
    wsP.Range("A1").Select
    DoEvents
    Err.Clear
    On Error GoTo FalhaImp

    destTmp = Environ$("TEMP") & "\Cecilia_graficos_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf"
    pdfOk = False
    On Error Resume Next
    Kill destTmp
    Err.Clear
    wsP.ExportAsFixedFormat Type:=xlTypePDF, Filename:=destTmp, Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    If Err.Number = 0 Then
        If Len(Dir$(destTmp)) > 0 Then pdfOk = True
    End If
    Err.Clear
    On Error GoTo FalhaImp

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error Resume Next
    Application.Calculation = calcAnt
    Err.Clear
    wsP.Activate
    If pdfOk Then
        If Application.Visible And Application.UserControl Then
            On Error Resume Next
            ThisWorkbook.FollowHyperlink destTmp
            Err.Clear
        End If
    ElseIf Application.Visible And Application.UserControl Then
        wsP.PrintPreview
    End If
    Exit Sub

FalhaImp:
    errN = Err.Number
    errD = Err.Description
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error Resume Next
    Application.StatusBar = False
    Application.Calculation = calcAnt
    If n > 0 Then
        wsP.Activate
        wsP.Range("A1").Select
        If Application.Visible And Application.UserControl Then wsP.PrintPreview
    ElseIf Not wsOld Is Nothing Then
        wsOld.Activate
        Aviso "Nao deu para montar a impressao. " & CStr(errN) & " " & errD
    End If
End Sub

Private Function GraficoNaLinha(ws As Worksheet, linha As Long) As ChartObject
    Dim co As ChartObject
    Dim r As Long
    Set GraficoNaLinha = Nothing
    If ws Is Nothing Or linha <= 0 Then Exit Function
    On Error Resume Next
    For Each co In ws.ChartObjects
        r = 0
        r = co.TopLeftCell.Row
        If r = linha Then
            Set GraficoNaLinha = co
            Exit Function
        End If
    Next co
End Function

Private Function ColocarGraficoNaFolha(co As ChartObject, wsP As Worksheet, _
    esquerda As Double, topo As Double, largura As Double, altura As Double, fontePts As Long) As Boolean
    Dim nAntes As Long
    Dim nShpAntes As Long
    Dim iShp As Long
    Dim novo As ChartObject
    Dim wsSrc As Worksheet
    Dim su As Boolean
    ColocarGraficoNaFolha = False
    On Error Resume Next
    Set wsSrc = co.Parent
    su = Application.ScreenUpdating
    nAntes = wsP.ChartObjects.Count
    nShpAntes = wsP.Shapes.Count
    co.Copy
    wsP.Paste
    If wsP.ChartObjects.Count <= nAntes Then
        Application.ScreenUpdating = True
        If Not wsSrc Is Nothing Then
            wsSrc.Activate
            co.Activate
            DoEvents
        End If
        co.Copy
        wsP.Activate
        wsP.Paste
        Application.ScreenUpdating = su
    End If
    If wsP.ChartObjects.Count <= nAntes Then
        For iShp = wsP.Shapes.Count To nShpAntes + 1 Step -1
            wsP.Shapes(iShp).Delete
        Next iShp
        Application.CutCopyMode = False
        Application.ScreenUpdating = su
        Err.Clear
        Exit Function
    End If
    Set novo = wsP.ChartObjects(wsP.ChartObjects.Count)
    novo.Left = esquerda
    novo.Top = topo
    novo.Width = largura
    novo.Height = altura
    AjustarGraficoImpresso novo, fontePts
    Application.CutCopyMode = False
    ColocarGraficoNaFolha = True
End Function

Public Sub AjustarAlturasDados()
    ' Compacta a lista de resultados (zoom ~82%, ~15 páginas).
    Dim ws As Worksheet
    Dim nomes As Variant
    Dim i As Long
    On Error Resume Next
    nomes = Array(ABA_DADOS_COMP, ABA_DADOS_SEL)
    For i = LBound(nomes) To UBound(nomes)
        Set ws = Folha(CStr(nomes(i)))
        If Not ws Is Nothing Then AjustarAlturasNaAba ws
    Next i
    On Error GoTo 0
End Sub

Private Sub AjustarAlturasNaAba(ws As Worksheet)
    Dim last As Long, r As Long
    If ws Is Nothing Then Exit Sub
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If last < 4 Then Exit Sub
    Application.ScreenUpdating = False
    On Error Resume Next
    ws.ResetAllPageBreaks
    Err.Clear
    On Error GoTo 0
    For r = 4 To last
        If EhSeparadorDados(ws, r) Then
            ws.Rows(r).RowHeight = 6
            GoTo ProxAlt
        End If
        On Error Resume Next
        ws.Rows(r).AutoFit
        Err.Clear
        On Error GoTo 0
        ' Aceita encolher (lista ~15 págs). Só limita extremos.
        If ws.Rows(r).RowHeight < 12 Then ws.Rows(r).RowHeight = 12
        If ws.Rows(r).RowHeight > 72 Then ws.Rows(r).RowHeight = 72
        If EhTituloExameDados(ws, r) Or EhTituloGrupoDados(ws, r) Then
            If ws.Rows(r).RowHeight < 16 Then ws.Rows(r).RowHeight = 16
        End If
ProxAlt:
    Next r
    ConfigurarZoomDados ws
    Application.ScreenUpdating = True
End Sub

Public Sub ConfigurarZoomDados(ws As Worksheet)
    ' Zoom fixo 82%: cabe na largura A4 paisagem (~15 páginas).
    On Error Resume Next
    Application.PrintCommunication = False
    Err.Clear
    ws.PageSetup.FitToPagesWide = False
    ws.PageSetup.FitToPagesTall = False
    ws.PageSetup.Zoom = 82
    Application.PrintCommunication = True
    Err.Clear
    ws.PageSetup.FitToPagesWide = False
    ws.PageSetup.FitToPagesTall = False
    ws.PageSetup.Zoom = 82
    Err.Clear
End Sub

Private Function EhTituloExameDados(ws As Worksheet, r As Long) As Boolean
    Dim cor As Long
    EhTituloExameDados = False
    On Error Resume Next
    If Len(Trim$(CStr(ws.Cells(r, 1).Value))) = 0 Then Exit Function
    cor = ws.Cells(r, 1).Interior.Color
    If cor = RGB(214, 234, 248) Or cor = RGB(217, 234, 247) Then
        EhTituloExameDados = True
    End If
End Function

Private Function EhTituloGrupoDados(ws As Worksheet, r As Long) As Boolean
    Dim cor As Long
    EhTituloGrupoDados = False
    On Error Resume Next
    If Len(Trim$(CStr(ws.Cells(r, 1).Value))) = 0 Then Exit Function
    cor = ws.Cells(r, 1).Interior.Color
    If cor = RGB(212, 230, 241) Then
        EhTituloGrupoDados = True
    End If
End Function

Private Function EhSeparadorDados(ws As Worksheet, r As Long) As Boolean
    EhSeparadorDados = False
    On Error Resume Next
    If Len(Trim$(CStr(ws.Cells(r, 1).Value))) > 0 Then Exit Function
    If Len(Trim$(CStr(ws.Cells(r, 2).Value))) > 0 Then Exit Function
    If Len(Trim$(CStr(ws.Cells(r, 6).Value))) > 0 Then Exit Function
    EhSeparadorDados = True
End Function

Private Function ColarTabelaUnicos(wsP As Worksheet, wsE As Worksheet, todos As Boolean, _
    nGraf As Long, ByRef r As Long, porPagina As Long) As Long
    Dim wsU As Worksheet
    Dim lastU As Long, i As Long, nLin As Long
    Dim er As Long
    Dim inclui As Boolean
    Dim zebra As Boolean
    Dim rng As Range
    Dim rCab As Long
    Dim rUlt As Long

    ColarTabelaUnicos = 0
    Set wsU = Folha(ABA_UNICOS)
    If wsU Is Nothing Then Exit Function
    lastU = wsU.Cells(wsU.Rows.Count, 1).End(xlUp).Row
    If lastU < 2 Then Exit Function

    nLin = 0
    For i = 2 To lastU
        inclui = todos
        If Not inclui Then
            er = 0
            On Error Resume Next
            er = CLng(wsU.Cells(i, 7).Value)
            Err.Clear
            On Error GoTo 0
            If er > 0 Then inclui = CelulaMarcada(wsE.Cells(er, 2).Value)
        End If
        If inclui Then nLin = nLin + 1
    Next i
    If nLin = 0 Then Exit Function

    If nGraf > 0 Then
        On Error Resume Next
        wsP.Rows(r).PageBreak = -4135
        Err.Clear
        On Error GoTo 0
    End If

    On Error Resume Next
    wsP.Range(wsP.Cells(r, 1), wsP.Cells(r, 6)).UnMerge
    Err.Clear
    On Error GoTo 0
    wsP.Range(wsP.Cells(r, 1), wsP.Cells(r, 6)).Merge
    With wsP.Cells(r, 1)
        .Value = "Exames com um unico resultado (sem grafico)"
        .Font.Bold = True
        .Font.Size = 12
        .Font.Name = "Calibri"
        .Font.Color = RGB(31, 78, 121)
        .Interior.Color = RGB(217, 234, 247)
        .VerticalAlignment = xlCenter
    End With
    wsP.Rows(r).RowHeight = 22
    r = r + 1

    On Error Resume Next
    wsP.Range(wsP.Cells(r, 1), wsP.Cells(r, 6)).UnMerge
    Err.Clear
    On Error GoTo 0
    wsP.Range(wsP.Cells(r, 1), wsP.Cells(r, 6)).Merge
    With wsP.Cells(r, 1)
        .Value = "So ha um ponto no laudo (inclui as dosagens pontuais). A evolucao aparece quando houver uma nova coleta."
        .Font.Size = 9
        .Font.Name = "Calibri"
        .Font.Italic = True
        .Font.Color = RGB(90, 90, 90)
        .Interior.Color = RGB(255, 255, 255)
    End With
    wsP.Rows(r).RowHeight = 18
    r = r + 1

    rCab = r
    wsP.Cells(r, 1).Value = "Exame"
    wsP.Cells(r, 2).Value = "Data"
    wsP.Cells(r, 3).Value = "Resultado"
    wsP.Cells(r, 4).Value = "Unidade"
    wsP.Cells(r, 5).Value = "Faixa do laudo"
    wsP.Cells(r, 6).Value = "Arquivo"
    Set rng = wsP.Range(wsP.Cells(r, 1), wsP.Cells(r, 6))
    With rng
        .Font.Bold = True
        .Font.Size = 9
        .Font.Name = "Calibri"
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    wsP.Rows(r).RowHeight = 18
    r = r + 1

    zebra = False
    For i = 2 To lastU
        inclui = todos
        If Not inclui Then
            er = 0
            On Error Resume Next
            er = CLng(wsU.Cells(i, 7).Value)
            Err.Clear
            On Error GoTo 0
            If er > 0 Then inclui = CelulaMarcada(wsE.Cells(er, 2).Value)
        End If
        If Not inclui Then GoTo ProxU

        wsP.Cells(r, 1).Value = Trim$(CStr(wsU.Cells(i, 1).Value))
        If IsDate(wsU.Cells(i, 2).Value) Then
            wsP.Cells(r, 2).Value = Format$(CDate(wsU.Cells(i, 2).Value), "dd/mm/yyyy")
        Else
            wsP.Cells(r, 2).Value = Trim$(CStr(wsU.Cells(i, 2).Value))
        End If
        wsP.Cells(r, 3).Value = Trim$(CStr(wsU.Cells(i, 3).Value))
        wsP.Cells(r, 4).Value = Trim$(CStr(wsU.Cells(i, 4).Value))
        wsP.Cells(r, 5).Value = Trim$(CStr(wsU.Cells(i, 5).Value))
        wsP.Cells(r, 6).Value = Trim$(CStr(wsU.Cells(i, 6).Value))

        Set rng = wsP.Range(wsP.Cells(r, 1), wsP.Cells(r, 6))
        With rng
            .Font.Size = 9
            .Font.Name = "Calibri"
            .Font.Color = RGB(40, 40, 40)
            .WrapText = True
            .VerticalAlignment = xlCenter
            If zebra Then
                .Interior.Color = RGB(217, 234, 247)
            Else
                .Interior.Color = RGB(255, 255, 255)
            End If
        End With
        wsP.Cells(r, 2).HorizontalAlignment = xlCenter
        wsP.Cells(r, 3).HorizontalAlignment = xlCenter
        wsP.Cells(r, 4).HorizontalAlignment = xlCenter
        On Error Resume Next
        wsP.Rows(r).AutoFit
        Err.Clear
        On Error GoTo 0
        If wsP.Rows(r).RowHeight < 16 Then wsP.Rows(r).RowHeight = 16
        If wsP.Rows(r).RowHeight > 48 Then wsP.Rows(r).RowHeight = 48

        zebra = Not zebra
        rUlt = r
        r = r + 1
        ColarTabelaUnicos = ColarTabelaUnicos + 1
ProxU:
    Next i

    If rUlt >= rCab Then
        AplicarGradeTabela wsP.Range(wsP.Cells(rCab, 1), wsP.Cells(rUlt, 6))
    End If
    If porPagina < 0 Then ColarTabelaUnicos = ColarTabelaUnicos
End Function

Private Sub AplicarGradeTabela(rng As Range)
    On Error Resume Next
    With rng.Borders
        .LineStyle = 1   ' xlContinuous
        .Weight = 2      ' xlThin
        .Color = RGB(140, 160, 180)
    End With
    With rng.Borders(8)  ' xlEdgeTop
        .LineStyle = 1
        .Weight = 2
        .Color = RGB(31, 78, 121)
    End With
    With rng.Borders(9)  ' xlEdgeBottom
        .LineStyle = 1
        .Weight = 2
        .Color = RGB(31, 78, 121)
    End With
    Err.Clear
End Sub

Private Sub ConfigurarColunasImpressao(wsP As Worksheet, usableW As Double)
    Dim alvos(1 To 6) As Double
    Dim i As Long
    Dim soma As Double, fat As Double
    ' Proporções da tabela; a soma final fica abaixo da largura útil
    alvos(1) = 132
    alvos(2) = 54
    alvos(3) = 68
    alvos(4) = 44
    alvos(5) = 94
    alvos(6) = 110
    soma = 0
    For i = 1 To 6
        soma = soma + alvos(i)
    Next i
    fat = (usableW - 16) / soma
    If fat > 1 Then fat = 1
    If fat < 0.5 Then fat = 0.5
    For i = 1 To 6
        AjustarLarguraColuna wsP, i, alvos(i) * fat
    Next i
End Sub

Private Function SomaLargurasColunas(wsP As Worksheet, c1 As Long, c2 As Long) As Double
    Dim c As Long
    SomaLargurasColunas = 0
    For c = c1 To c2
        SomaLargurasColunas = SomaLargurasColunas + wsP.Columns(c).Width
    Next c
End Function

Private Sub CaberNaLarguraDaPagina(wsP As Worksheet, usableW As Double)
    Dim lim As Double
    Dim co As ChartObject
    Dim soma As Double
    Dim fat As Double
    Dim i As Long
    Dim alvo As Double
    On Error Resume Next
    lim = usableW - 10
    If lim < 300 Then lim = 300

    soma = SomaLargurasColunas(wsP, 1, 6)
    If soma > lim Then
        fat = lim / soma
        For i = 1 To 6
            alvo = wsP.Columns(i).Width * fat
            AjustarLarguraColuna wsP, i, alvo
        Next i
    End If

    lim = SomaLargurasColunas(wsP, 1, 6) - 4
    For Each co In wsP.ChartObjects
        If co.Left < 0 Then co.Left = 2
        If co.Left + co.Width > lim Then
            If lim - co.Left > 50 Then co.Width = lim - co.Left
        End If
    Next co
    Err.Clear
End Sub

Private Sub AjustarGraficoImpresso(co As ChartObject, fontePts As Long)
    Dim ch As Chart
    On Error Resume Next
    Set ch = co.Chart
    ch.Axes(1).HasTitle = False
    ch.Axes(1).TickLabels.Font.Name = "Calibri"
    ch.Axes(1).TickLabels.Font.Size = fontePts
    ch.Axes(2).TickLabels.Font.Name = "Calibri"
    ch.Axes(2).TickLabels.Font.Size = fontePts
    If ch.Axes(2).HasTitle Then ch.Axes(2).AxisTitle.Font.Size = fontePts + 1
    AjustarLegendaGrafico ch, fontePts
End Sub

Private Sub AjustarLegendaGrafico(ch As Chart, fontePts As Long)
    Dim i As Long
    Dim nSer As Long
    Dim nNomes As Long
    Dim nm As String
    On Error Resume Next
    nSer = ch.SeriesCollection.Count
    nNomes = 0
    For i = 1 To nSer
        nm = CStr(ch.SeriesCollection(i).Name)
        If InStr(1, nm, "_base", vbTextCompare) = 0 And InStr(1, nm, "piso", vbTextCompare) = 0 Then
            nNomes = nNomes + 1
        End If
    Next i
    If nNomes < 2 Then
        ch.HasLegend = False
        Exit Sub
    End If
    ch.HasLegend = False
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom
    ch.Legend.IncludeInLayout = True
    If fontePts > 0 Then
        ch.Legend.Font.Name = "Calibri"
        ch.Legend.Font.Size = fontePts
    End If
    For i = ch.Legend.LegendEntries.Count To 1 Step -1
        nm = ""
        If i <= nSer Then nm = CStr(ch.SeriesCollection(i).Name)
        If InStr(1, nm, "_base", vbTextCompare) > 0 Or InStr(1, nm, "piso", vbTextCompare) > 0 Then
            ch.Legend.LegendEntries(i).Delete
        End If
    Next i
End Sub

Private Sub AjustarLarguraColuna(ws As Worksheet, col As Long, alvoPts As Double)
    Dim lo As Double, hi As Double, mid As Double
    Dim k As Long
    Dim best As Double
    On Error Resume Next
    If alvoPts < 20 Then alvoPts = 20
    lo = 1
    hi = 120
    best = lo
    For k = 1 To 14
        mid = (lo + hi) / 2
        ws.Columns(col).ColumnWidth = mid
        If ws.Columns(col).Width <= alvoPts Then
            best = mid
            lo = mid
        Else
            hi = mid
        End If
    Next k
    ws.Columns(col).ColumnWidth = best
End Sub

Public Sub GarantirNomesGraficos(wsG As Worksheet)
    Dim wsP As Worksheet
    Dim wsSrc As Worksheet
    Dim co As ChartObject
    Dim k As Long
    Dim r As Long
    Dim nm As String

    Set wsP = FolhaPosicoes()
    If Len(CStr(wsP.Cells(2, 1).Value)) > 0 Then Exit Sub
    Set wsSrc = Folha(ABA_GRAF_COMP)
    If wsSrc Is Nothing Then Set wsSrc = wsG
    If wsSrc Is Nothing Then Exit Sub
    If wsSrc.ChartObjects.Count = 0 Then Exit Sub

    wsP.Cells(1, 1).Value = "nome"
    wsP.Cells(1, 2).Value = "linha"
    wsP.Cells(1, 3).Value = "largura"
    wsP.Cells(1, 4).Value = "altura"

    k = 2
    For Each co In wsSrc.ChartObjects
        r = co.TopLeftCell.Row
        nm = "G" & CStr(r)
        On Error Resume Next
        co.Name = nm
        On Error GoTo 0
        wsP.Cells(k, 1).Value = co.Name
        wsP.Cells(k, 2).Value = r
        wsP.Cells(k, 3).Value = co.Width
        wsP.Cells(k, 4).Value = co.Height
        k = k + 1
    Next co
    If Not wsG Is Nothing Then
        If Not wsG Is wsSrc Then
            For Each co In wsG.ChartObjects
                r = 0
                On Error Resume Next
                r = co.TopLeftCell.Row
                If r > 0 Then co.Name = "G" & CStr(r)
                On Error GoTo 0
            Next co
        End If
    End If
    If k > 3 Then
        wsP.Range("A1:D" & CStr(k - 1)).Sort Key1:=wsP.Range("B2"), Order1:=1, Header:=1
    End If
End Sub

Private Sub AplicarGraficos(wsG As Worksheet, wsM As Worksheet, wsE As Worksheet, last As Long)
    Dim wsP As Worksheet
    Dim lastP As Long, i As Long
    Dim nm As String
    Dim r As Long
    Dim w As Double, h As Double
    Dim show As Boolean
    Dim co As ChartObject

    Set wsP = FolhaPosicoes()
    lastP = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row
    If lastP < 2 Then Exit Sub

    For i = 2 To lastP
        nm = CStr(wsP.Cells(i, 1).Value)
        If Len(nm) = 0 Then GoTo Prox
        r = CLng(wsP.Cells(i, 2).Value)
        w = CDbl(wsP.Cells(i, 3).Value)
        h = CDbl(wsP.Cells(i, 4).Value)
        show = GraficoMarcado(wsM, wsE, last, r)

        On Error Resume Next
        Set co = Nothing
        Set co = wsG.ChartObjects(nm)
        On Error GoTo 0
        If co Is Nothing Then GoTo Prox

        On Error Resume Next
        co.Placement = xlMoveAndSize
        If show Then
            co.Visible = True
            co.ShapeRange.Visible = msoTrueVal
            co.Left = wsG.Cells(r, 1).Left
            co.Top = wsG.Cells(r, 1).Top
            If w > 20 Then co.Width = w
            If h > 20 Then co.Height = h
        Else
            co.Visible = False
            co.ShapeRange.Visible = msoFalseVal
        End If
        On Error GoTo 0
Prox:
    Next i
End Sub

Private Function FolhaImpressao(porPagina As Long) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Set ws = ThisWorkbook.Worksheets("Impressao")
    If Not ws Is Nothing Then ws.Delete
    Application.DisplayAlerts = True
    Err.Clear
    On Error GoTo 0
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = "Impressao"
    ws.Tab.Color = RGB(52, 73, 94)
    On Error Resume Next
    Application.PrintCommunication = False
    Err.Clear
    On Error GoTo 0
    With ws.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(1)
        .HeaderMargin = Application.CentimetersToPoints(0.4)
        .FooterMargin = Application.CentimetersToPoints(0.4)
        .CenterHorizontally = False
        .PrintGridlines = False
        .CenterHeader = "Cecilia Maria Albergaria Silva  -  graficos"
        .CenterFooter = CStr(porPagina) & " por folha  |  &P / &N"
        .PrintTitleRows = ""
    End With
    On Error Resume Next
    Application.PrintCommunication = True
    ' Reaplica com comunicação ligada (senão o Excel ignora margens/zoom)
    With ws.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(1)
    End With
    Err.Clear
    Set FolhaImpressao = ws
End Function

Private Function FolhaPosicoes() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("GrafObj")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "GrafObj"
        ws.Visible = xlVeryHidden
    End If
    Set FolhaPosicoes = ws
End Function

Private Sub Aviso(msg As String)
    Dim p As String
    On Error Resume Next
    p = Environ$("TEMP") & "\cecilia_print_log.txt"
    Open p For Append As #1
    Print #1, Now & " " & msg
    Close #1
    If Application.UserControl Then MsgBox msg, vbExclamation
End Sub

Private Function GraficoMarcado(wsM As Worksheet, wsE As Worksheet, last As Long, r As Long) As Boolean
    Dim i As Long, er As Long, g1 As Long, g2 As Long
    GraficoMarcado = False
    If r <= 0 Then Exit Function
    For i = 2 To last
        g1 = 0: g2 = 0
        If Not IsEmpty(wsM.Cells(i, 4).Value) Then g1 = CLng(wsM.Cells(i, 4).Value)
        If Not IsEmpty(wsM.Cells(i, 5).Value) Then g2 = CLng(wsM.Cells(i, 5).Value)
        If g1 > 0 And r >= g1 And r <= g2 Then
            er = CLng(wsM.Cells(i, 1).Value)
            GraficoMarcado = CelulaMarcada(wsE.Cells(er, 2).Value)
            Exit Function
        End If
    Next i
End Function

Private Function CelulaMarcada(v As Variant) As Boolean
    CelulaMarcada = False
    If VarType(v) = vbBoolean Then
        CelulaMarcada = v
    ElseIf IsNumeric(v) Then
        CelulaMarcada = (CDbl(v) <> 0)
    ElseIf UCase$(CStr(v)) = "TRUE" Or CStr(v) = "Sim" Then
        CelulaMarcada = True
    End If
End Function

Public Sub EsconderSeriesInvisiveis()
    Dim ws As Worksheet
    Dim co As ChartObject
    Dim nomes As Variant
    Dim i As Long
    On Error Resume Next
    nomes = Array(ABA_GRAF_COMP, ABA_GRAF_SEL)
    For i = LBound(nomes) To UBound(nomes)
        Set ws = Folha(CStr(nomes(i)))
        If ws Is Nothing Then GoTo ProxWs
        For Each co In ws.ChartObjects
            AjustarLegendaGrafico co.Chart, 11
        Next co
ProxWs:
    Next i
End Sub
