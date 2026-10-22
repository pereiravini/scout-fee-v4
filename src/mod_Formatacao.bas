Attribute VB_Name = "Módulo1"
Sub format_Prod()
Attribute format_Prod.VB_ProcData.VB_Invoke_Func = "j\n14"
'
' formatar Macro
'
' Atalho do teclado: Ctrl+j
'
    Sheets("Produtividade").Select
    Range("A1:R5").Select
    Selection.UnMerge
    Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:C").Select
    Selection.Delete Shift:=xlToLeft
    Columns("D:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:J").Select
    Selection.Delete Shift:=xlToLeft
    Columns("F:G").Select
    Selection.Delete Shift:=xlToLeft
    Range("I6").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "Status"
    Range("I6").Select
    Selection.Font.Bold = False
    Selection.Font.Bold = True
    Columns("A:I").Select
    Columns("A:I").EntireColumn.AutoFit
    Range("A6").Select
    Selection.AutoFilter
    Columns("A:I").Select
    Columns("A:I").EntireColumn.AutoFit
    Range("A6").Select
    Sheets("Produtividade").Select
    Range("A6").Select
    ActiveWorkbook.Worksheets("Produtividade").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Produtividade").AutoFilter.Sort.SortFields.Add Key _
        :=Range("A6:A314"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Produtividade").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Produtividade").AutoFilter.Sort.SortFields.Clear
    Sheets("Produtividade").Select
    Sheets("Suporte").Visible = True
    Sheets("Suporte").Select
    Range("A10:B14").Select
    Selection.Copy
    Sheets("Produtividade").Select
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    
'    Sheets("Suporte").Select
'    ActiveWindow.SelectedSheets.Visible = False
    
    Dim ws As Worksheet
    Set ws = Sheets("Produtividade")

    On Error Resume Next
    ws.Shapes("btn_FormatProd").Delete
    On Error GoTo 0
    
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim shp As Shape

    Set wsOrigem = Sheets("Suporte")
    Set wsDestino = Sheets("Produtividade")

' Botão AnálisePP em C2
    wsOrigem.Shapes("btn_AnalisePP").Copy
    wsDestino.Paste
    Set shp = wsDestino.Shapes(wsDestino.Shapes.Count)
    With shp
        .Top = wsDestino.Range("C2").Top
        .Left = wsDestino.Range("C2").Left
        .Placement = xlMove
    End With

' Botão Limpar em C4
    wsOrigem.Shapes("btn_Limpar").Copy
    wsDestino.Paste
    Set shp = wsDestino.Shapes(wsDestino.Shapes.Count)
    With shp
        .Top = wsDestino.Range("C4").Top
        .Left = wsDestino.Range("C4").Left
        .Placement = xlMove
    End With
    Sheets("Suporte").Select
    ActiveWindow.SelectedSheets.Visible = False
    
    ActiveWorkbook.Worksheets("Produtividade").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Produtividade").AutoFilter.Sort.SortFields.Add2 Key _
        :=Range("B6:B12"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Produtividade").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Produtividade").AutoFilter.Sort.SortFields.Clear
       
    Sheets("Produtividade").Select
    Range("A6").Select
    
End Sub
