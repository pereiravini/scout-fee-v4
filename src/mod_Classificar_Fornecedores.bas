Attribute VB_Name = "Módulo10"
Sub classificarFornecedores()

    Dim ws As Worksheet
    Dim ultimaLinha As Long

    Set ws = ThisWorkbook.Worksheets("Registros_Fee")
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    With ws.Range("B7:B" & ultimaLinha)
        .Value = .Value
    End With

    'Atualiza o filtro para incluir todos os dados
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range("A6:E" & ultimaLinha).AutoFilter

    'Ordena a tabela inteira pelo Fornecedor (coluna B)
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("B7:B" & ultimaLinha), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal

        .SetRange ws.Range("A6:E" & ultimaLinha)
        .Header = xlYes
        .Apply
    End With

End Sub

