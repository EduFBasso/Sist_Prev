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
    Dim ws As Worksheet
    Dim ultima As Long
    Dim i As Long
    
    On Error Resume Next
    Set ws = Sheets("Config_Regras")
    On Error GoTo 0
    
    If ws Is Nothing Then
        GetParametro = 0
        Exit Function
    End If
    
    ' Procura o parâmetro na planilha (assume coluna 1 = nome, coluna 2 = valor)
    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To ultima  ' Começa na linha 2 (pula cabeçalho)
        If Trim(UCase(ws.Cells(i, 1).Value)) = Trim(UCase(nome)) Then
            GetParametro = ws.Cells(i, 2).Value
            Exit Function
        End If
    Next i
    
    ' Se não encontrou, retorna 0 (mais seguro que erro)
    GetParametro = 0
End Function
