Attribute VB_Name = "Módulo4"
Sub limpar()
Attribute limpar.VB_ProcData.VB_Invoke_Func = " \n14"
'
' limpar Macro
'
'
    Sheets("Produtividade").Select
    If ActiveSheet.AutoFilterMode Then
     ActiveSheet.AutoFilterMode = False
    End If
    Columns("H:K").Select
    ActiveWindow.SmallScroll ToRight:=1
    Selection.EntireColumn.Hidden = False
    ActiveWindow.ScrollColumn = 1
    Range("A6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Sheets("Registros_Fee").Select
    Cells.Select
    Selection.ClearContents
    Range("F2:H4").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    Sheets("Suporte").Visible = True
    Sheets("Suporte").Select
    Range("A10:B14").Select
    Selection.Copy
    Sheets("Registros_Fee").Select
    Range("A1:B5").Select
    ActiveSheet.Paste
    Range("A1").Select
    Sheets("Suporte").Visible = False
    Sheets("Produtividade").Select
    
    Dim ws As Worksheet

    Set ws = Sheets("Produtividade")

    On Error Resume Next
    ws.Shapes("btn_AnalisePP").Delete
    ws.Shapes("btn_Limpar").Delete
    On Error GoTo 0
    
    Set ws = Sheets("Registros_Fee")

    On Error Resume Next
    ws.Shapes("btn_FeeArq").Delete
    ws.Shapes("btn_ClassForn").Delete
    On Error GoTo 0
    
    Sheets("Fee Por Arquivo").Select
    Range("A7").Select
    ActiveCell.Range("A1:B1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("A7").Select
    
' Botão Formatar em C3
    
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim shp As Shape

    Set wsOrigem = Sheets("Suporte")
    Set wsDestino = Sheets("Produtividade")

    wsOrigem.Shapes("btn_FormatProd").Copy
    wsDestino.Paste
    Set shp = wsDestino.Shapes(wsDestino.Shapes.Count)
    With shp
        .Top = wsDestino.Range("C3").Top
        .Left = wsDestino.Range("C3").Left
        .Placement = xlMove
    End With
       
    Sheets("Produtividade").Select
    Range("A6").Select

End Sub
