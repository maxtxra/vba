Sub ManterColunasEspecificas()
    Dim ws As Worksheet
    Dim col As Long
    
    ' Defina a planilha em que você deseja manter as colunas
    Set ws = ThisWorkbook.Sheets("NomeDaPlanilha") ' Substitua "NomeDaPlanilha" pelo nome correto da planilha
    
    Application.ScreenUpdating = False
    
    ' Percorra todas as colunas e exclua se não estiver na lista de colunas para manter
    For col = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column To 1 Step -1
        If Not (col = 1 Or col = 4 Or col = 5 Or col = 8 Or col = 10 Or col = 11 Or col = 25 Or col = 38 Or col = 49) Then
            ws.Columns(col).Delete
        End If
    Next col
    
    Application.ScreenUpdating = True
End Sub
