Attribute VB_Name = "Módulo3"
Sub analise_key_pp()
Attribute analise_key_pp.VB_ProcData.VB_Invoke_Func = " \n14"
'
' analise_key_pp Macro
'

'
    Sheets("Produtividade").Select
    Range("AZ1") = InputBox("Fee aéreo/seguro viagem já foi apurado?")
    If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
    End If
    
    Sheets("Registros_Fee").Select
    Range("B6").Select
    ActiveCell.FormulaR1C1 = "Fornecedor"
    Range("B7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Columns("C:G").Select
    Selection.Delete Shift:=xlToLeft
    Range("B7").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(INDEX(Produtividade!C,MATCH(Registros_Fee!RC[-1],Produtividade!C[1],0)),""NÃO LOCALIZADO"")"
    Range("B7").Select
    Selection.Copy
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Cells.Select
    Cells.EntireColumn.AutoFit
    
    Range("A7").Select

    Sheets("Produtividade").Select

    Range("J6").Select
    ActiveCell.FormulaR1C1 = "Fee apurado?"
    Range("K6").Select
    ActiveCell.FormulaR1C1 = "Qtd_Apurações"
    Range("L6").Select
    ActiveCell.FormulaR1C1 = "Análise Final"
    Range("J6:L6").Select
    Selection.Font.Bold = True
    Range("J7").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(VLOOKUP(RC[-7],Registros_Fee!C[-9]:C[-8],1,0)<>""#N/D"",""SIM"",""""),""NÃO"")"
    Range("J7").Select
    Selection.Copy
    ActiveCell.Offset(0, -2).Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("K7").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1]=""NÃO"",0,COUNTIF(Registros_Fee!C[-10],RC[-8]))"
    Range("K7").Select
    Selection.Copy
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("L7").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-2]=""SIM"",LEFT(RC[-11],5)=""CUSTO"",SUMIF(Registros_Fee!C1,RC[-9],Registros_Fee!C4)>0),Suporte!R2C1, IF(AND(RC[-2]=""NÃO"",LEFT(R1C52,1)=""s"",OR(RC[-8]=""AÉREO"",RC[-8]=""SEGURO VIAGEM"")),""FEE DEVIDO - JÁ APURADO"", IF(AND(RC[-2]=""SIM"",R1C52=""SIM"",OR(RC[-8]=""AÉREO"",RC[-8]=""SEGURO VIAGEM"")),Suporte!R5C1, IF(AND(RC[-2]=""NÃO"",LEFT(RC[-11],5)<>" & _
        """CUSTO""),Suporte!R3C1, IF(AND(RC[-2]=""SIM"",SUMIF(Registros_Fee!C1,RC[-9],Registros_Fee!C4)=0,LEFT(RC[-11],5)<>""CUSTO""),Suporte!R6C1, IF(AND(RC[-1]>1,LEFT(RC[-10],2)<>""MM""),Suporte!R1C1, IF(AND(RC[-8]<>""AÉREO"",RC[-6]<>IFERROR(VLOOKUP(RC[-9],Registros_Fee!C1:C4,3,0),0),LEFT(RC[-11],5)<>""CUSTO""),Suporte!R7C1, ""FEE DEVIDO"")))))))" & _
        ""
    Range("L7").Select
    Selection.Copy
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Columns("J:L").Select
    Columns("J:L").EntireColumn.AutoFit
    Columns("L:L").Select
    Selection.FormatConditions.Add Type:=xlTextString, String:="problema", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:="devido", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Sheets("Registros_Fee").Select
    Range("E6").Select
    ActiveCell.FormulaR1C1 = "Status"
    Range("E6").Select
    Selection.Font.Bold = True
    Range("E7").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-3]=""NÃO LOCALIZADO"",RC[-3],VLOOKUP(RC[-4],Produtividade!C[-2]:C[4],7,0))"
    Range("E7").Select
    Selection.Copy
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
'   Quadro de resumos em Registros_Fee
    
    Sheets("Suporte").Visible = True
    Sheets("Suporte").Select
    Range("E11:G13").Select
    Selection.Copy
    Sheets("Registros_Fee").Select
    Range("F2").Select
    ActiveSheet.Paste
    Range("F4").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(C[-2])"
    Range("H4").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-2]-RC[-1])"
    Columns("F:F").ColumnWidth = 14.78
    Columns("G:G").ColumnWidth = 14.78
    Columns("H:H").ColumnWidth = 14.78
    Range("A7").Select
    Sheets("Suporte").Visible = False
    
    Sheets("Produtividade").Select
    Columns("J:K").Select
    Selection.EntireColumn.Hidden = True
    Range("A6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    
'   Botão Fee Por Arquivo em C2 e Classificar Fornecedores em C4
    
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim shp As Shape

    Set wsOrigem = Sheets("Suporte")
    Set wsDestino = Sheets("Registros_Fee")

    wsOrigem.Shapes("btn_FeeArq").Copy
    wsDestino.Paste
    Set shp = wsDestino.Shapes(wsDestino.Shapes.Count)
    With shp
        .Top = wsDestino.Range("C2").Top
        .Left = wsDestino.Range("C2").Left
        .Placement = xlMove
    End With
    
    wsOrigem.Shapes("btn_ClassForn").Copy
    wsDestino.Paste
    Set shp = wsDestino.Shapes(wsDestino.Shapes.Count)
    With shp
        .Top = wsDestino.Range("C4").Top
        .Left = wsDestino.Range("C4").Left
        .Placement = xlMove
    End With
       
    Sheets("Registros_Fee").Select
    
    Range("A6").Select
    
End Sub
