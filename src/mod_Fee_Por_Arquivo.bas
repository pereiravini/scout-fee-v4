Attribute VB_Name = "Módulo7"
Sub feePorArquivo()

    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim linha As Long
    Dim rloc As String
    Dim valor As Double
    
    With Sheets("Registros_Fee")
        .Range("A7", .Range("A7").End(xlDown)).Copy
    End With
    
    With Sheets("Fee Por Arquivo")
        .Range("A7").PasteSpecial Paste:=xlPasteValues
    End With
    
    With Sheets("Registros_Fee")
        .Range("D7", .Range("D7").End(xlDown)).Copy
    End With
    
    With Sheets("Fee Por Arquivo")
        .Range("B7").PasteSpecial Paste:=xlPasteValues
    End With

    Set ws = Sheets("Fee Por Arquivo")
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Application.ScreenUpdating = False

    'De baixo para cima para não pular linhas ao excluir
    For linha = ultimaLinha To 7 Step -1
        
        rloc = Trim(ws.Cells(linha, "A").Value)
        valor = Val(ws.Cells(linha, "B").Value)

        If rloc <> "" Then
            If Application.WorksheetFunction.CountIf(ws.Range("A2:A" & ultimaLinha), rloc) > 1 _
               And valor > 0 Then
               
                ws.Rows(linha).Delete
            End If
        End If

    Next linha

    Application.ScreenUpdating = True
    
    Sheets("Fee Por Arquivo").Select
    Range("A7").Select

End Sub

