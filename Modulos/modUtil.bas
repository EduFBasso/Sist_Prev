' ============================================
' modUtil - Funções utilitárias do sistema
' ============================================

Option Explicit

' Gera um novo ID incremental
Function NovoID(tabela As String, colunaID As Long) As Long
    Dim ws As Worksheet
    Dim ultima As Long
    
    Set ws = Sheets(tabela)
    ultima = ws.Cells(ws.Rows.count, colunaID).End(xlUp).Row
    
    If ultima < 2 Then
        NovoID = 1
    Else
        NovoID = ws.Cells(ultima, colunaID).Value + 1
    End If
End Function

' Converte texto vazio em zero
Function Nz(valor)
    If Trim(valor & "") = "" Then
        Nz = 0
    Else
        Nz = valor
    End If
End Function
