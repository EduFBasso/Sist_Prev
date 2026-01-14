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

' ============================================
' GetDadosEmpresa - Busca dados da empresa
' ============================================
' Retorna valor de configuração da planilha Config_Empresa
' Parâmetros:
'   campo - Nome do parâmetro (ex: "Razao_Social", "CNPJ", "Telefone")
' Retorna: String com o valor ou "" se não encontrado
' ============================================
Function GetDadosEmpresa(campo As String) As String
    Dim ws As Worksheet
    Dim i As Long
    Dim ultimaLinha As Long
    
    On Error Resume Next
    Set ws = Sheets("Config_Empresa")
    On Error GoTo 0
    
    ' Se planilha não existe, retornar vazio
    If ws Is Nothing Then
        GetDadosEmpresa = ""
        Exit Function
    End If
    
    ' Buscar campo na coluna A (comparação case-insensitive)
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To ultimaLinha
        If UCase(Trim(ws.Cells(i, 1).Value)) = UCase(Trim(campo)) Then
            GetDadosEmpresa = Trim(ws.Cells(i, 2).Value)
            Exit Function
        End If
    Next i
    
    ' Campo não encontrado
    GetDadosEmpresa = ""
End Function
