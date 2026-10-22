Attribute VB_Name = "Módulo14"
Sub atualizar_Prod()
Attribute atualizar_Prod.VB_ProcData.VB_Invoke_Func = "d\n14"
'
' formatar Macro
'
' Atalho do teclado: Ctrl+j
'

    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim ultimaDestino As Long
    
    
    Set ws = Sheets("Novos Registros")

    On Error Resume Next
    ws.Shapes("btn_AtualizarProd").Delete
    On Error GoTo 0


    Sheets("Novos Registros").Select
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
    Sheets("Novos Registros").Select
    Range("A6").Select
    Sheets("Suporte").Visible = True
    Sheets("Suporte").Select
    Range("A10:B14").Select
    Selection.Copy
    Sheets("Novos Registros").Select
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
    
    
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim shp As Shape

    Set wsOrigem = Sheets("Suporte")
    Set wsDestino = Sheets("Novos Registros")

    wsOrigem.Shapes("btn_AtualizarProd").Copy
    wsDestino.Paste
    Set shp = wsDestino.Shapes(wsDestino.Shapes.Count)
    With shp
        .Top = wsDestino.Range("C3").Top
        .Left = wsDestino.Range("C3").Left
        .Placement = xlMove
    End With
    
    
    Sheets("Suporte").Select
    ActiveWindow.SelectedSheets.Visible = False
       
    Sheets("Novos Registros").Select
    Range("A6").Select
    
    'Set ws = Sheets("Novos Registros") duplicado

    ultimaLinha = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row

    ' PROCV na coluna I
    ws.Range("I7:I" & ultimaLinha).FormulaR1C1 = _
        "=VLOOKUP(RC[-6],Produtividade!C[-6]:C,7,0)"
        
    ' Substitui fórmulas por valores
    ws.Range("I7:I" & ultimaLinha).Value = _
        ws.Range("I7:I" & ultimaLinha).Value

    ' filtra tudo que NÃO é erro
    ws.Range("A6:I" & ultimaLinha).AutoFilter _
        Field:=9, Criteria1:="<>#N/A"

    ' apaga linhas visíveis
    On Error Resume Next
    ws.Range("A7:A" & ultimaLinha).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    On Error GoTo 0

    ws.AutoFilterMode = False

    ' recalcula última linha
    ultimaLinha = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    
    ' quantidade de registros novos
    qtdNovos = ultimaLinha - 6

    ' limpa coluna STATUS
    If ultimaLinha >= 7 Then
        ws.Range("I7:I" & ultimaLinha).ClearContents
    End If
    
    ' mensagens
    If qtdNovos > 1 Then
    
        Sheets("Produtividade").Select
        Range("A6").Select
        If ActiveSheet.AutoFilterMode Then
        ActiveSheet.AutoFilterMode = False
        End If
        Sheets("Novos Registros").Select
        Range("A6").Select
        ActiveCell.Offset(1, 0).Range("A1").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        Sheets("Produtividade").Select
        Range("A6").Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(1, 0).Range("A1").Select
        ActiveSheet.Paste
        Sheets("Novos Registros").Select
        Range("A6").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.ClearContents
        Range("A6").Select
        Sheets("Produtividade").Select
    
        MsgBox qtdNovos & " novos registros encontrados e adicionados à aba Produtividade. Operação concluída com sucesso!", vbInformation
    
    ElseIf qtdNovos > 0 Then
    
    Sheets("Produtividade").Select
        Range("A6").Select
        If ActiveSheet.AutoFilterMode Then
        ActiveSheet.AutoFilterMode = False
        End If
        Sheets("Novos Registros").Select
        Range("A6").Select
        ActiveCell.Offset(1, 0).Range("A1").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Copy
        Sheets("Produtividade").Select
        Range("A6").Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(1, 0).Range("A1").Select
        ActiveSheet.Paste
        Sheets("Novos Registros").Select
        Range("A6").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.ClearContents
        Range("A6").Select
        Sheets("Produtividade").Select
    
        MsgBox qtdNovos & " novo registro encontrado e adicionado à aba Produtividade. Operação concluída com sucesso!", vbInformation
    
    Else
    
        Sheets("Novos Registros").Select
        Range("A6").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.ClearContents
        Range("A6").Select
        Sheets("Produtividade").Select
        
        MsgBox "Nenhum registro novo encontrado.", vbExclamation
    End If

    
    
End Sub

