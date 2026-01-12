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

' ============================================
' Nz - Retorna valor padrão se nulo/vazio
' ============================================
' Função compatível com Access VBA para Excel
' Retorna valorPadrao se o valor for NULL, Empty ou vazio
' Parâmetros:
'   valor - Valor a verificar
'   valorPadrao - Valor a retornar se nulo (padrão = 0)
' ============================================
Function Nz(valor As Variant, Optional valorPadrao As Variant = 0) As Variant
    If IsNull(valor) Or IsEmpty(valor) Or Trim(CStr(valor)) = "" Then
        Nz = valorPadrao
    Else
        Nz = valor
    End If
End Function
