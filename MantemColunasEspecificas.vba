Sub ManterColunasEspecificas()
    Dim ws As Worksheet
    Dim col As Long
    
    ' Define a aba planilha em que você deseja manter as colunas (não é o mesmo que o nomedoarquivo.xlsx -talvez)
    
    Set ws = ThisWorkbook.Sheets("Report") ' Substitua "Report" pelo nome correto da aba (se houver mais de uma) do excel
    
    Application.ScreenUpdating = False
    
    ' Procura todas as colunas (em ordem númerica e não em letras) que não devem ser deletadas e deleta todas as outras (alterar para seus números no if not)
    
    
    For col = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column To 1 Step -1
        If Not (col = 1 Or col = 4 Or col = 5 Or col = 8 Or col = 10 Or col = 11 Or col = 25 Or col = 38 Or col = 49) Then
            ws.Columns(col).Delete
        End If
    Next col
    
    Application.ScreenUpdating = True
End Sub

