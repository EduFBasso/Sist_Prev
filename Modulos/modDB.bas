' ============================================
' modDB - Acesso às planilhas (banco de dados)
' ============================================

Option Explicit

' Retorna a primeira linha disponível
Function ProximaLinha(ws As Worksheet) As Long
    ProximaLinha = ws.Cells(ws.Rows.count, 1).End(xlUp).Row + 1
End Function

' Lê um parâmetro da aba Config_Regras
Function GetParametro(nome As String) As Variant
    Dim tbl As ListObject
    Dim linha As ListRow
    
    Set tbl = Sheets("Config_Regras").ListObjects(1)
    
    For Each linha In tbl.ListRows
        If linha.Range(1, 1).Value = nome Then
            GetParametro = linha.Range(1, 2).Value
            Exit Function
        End If
    Next linha
    
    GetParametro = CVErr(xlErrNA)
End Function
