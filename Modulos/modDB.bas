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
' Grava ou atualiza um parâmetro na aba Config_Regras
Sub SetParametro(nome As String, valor As Variant, Optional descricao As String = "")
    Dim ws As Worksheet
    Dim ultima As Long
    Dim i As Long
    Dim encontrado As Boolean
    
    On Error Resume Next
    Set ws = Sheets("Config_Regras")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Planilha Config_Regras não encontrada!", vbCritical
        Exit Sub
    End If
    
    ' Procura o parâmetro existente
    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    encontrado = False
    
    For i = 2 To ultima  ' Começa na linha 2 (pula cabeçalho)
        If Trim(UCase(ws.Cells(i, 1).Value)) = Trim(UCase(nome)) Then
            ' Atualizar valor existente
            ws.Cells(i, 2).Value = valor
            ws.Cells(i, 4).Value = Date  ' Atualiza data de modificação (coluna 4)
            If descricao <> "" Then
                ws.Cells(i, 3).Value = descricao
            End If
            encontrado = True
            Exit For
        End If
    Next i
    
    ' Se não encontrou, adicionar novo parâmetro
    If Not encontrado Then
        Dim novaLinha As Long
        novaLinha = ultima + 1
        
        ws.Cells(novaLinha, 1).Value = nome
        ws.Cells(novaLinha, 2).Value = valor
        ws.Cells(novaLinha, 3).Value = descricao
        ws.Cells(novaLinha, 4).Value = Date
    End If
    
End Sub

' ============================================
' Configurar fatores de conversão de tempo especial
' ============================================
Sub ConfigurarFatoresEspeciais()
    ' Configura os 6 fatores de conversão conforme legislação
    ' Baseado na tabela oficial de conversão de tempo especial
    
    ' HOMEM - Fatores de conversão
    Call SetParametro("Conversao_Especial_15_H", 2.33, "Fator conversão especial 15 anos (Homem)")
    Call SetParametro("Conversao_Especial_20_H", 1.75, "Fator conversão especial 20 anos (Homem)")
    Call SetParametro("Conversao_Especial_25_H", 1.4, "Fator conversão especial 25 anos (Homem)")
    
    ' MULHER - Fatores de conversão
    Call SetParametro("Conversao_Especial_15_M", 2, "Fator conversão especial 15 anos (Mulher)")
    Call SetParametro("Conversao_Especial_20_M", 1.5, "Fator conversão especial 20 anos (Mulher)")
    Call SetParametro("Conversao_Especial_25_M", 1.2, "Fator conversão especial 25 anos (Mulher)")
    
    MsgBox "Fatores de conversão de tempo especial configurados com sucesso!" & vbCrLf & vbCrLf & _
           "HOMEM: 15 anos (×2.33) | 20 anos (×1.75) | 25 anos (×1.40)" & vbCrLf & _
           "MULHER: 15 anos (×2.00) | 20 anos (×1.50) | 25 anos (×1.20)", _
           vbInformation, "Configuração Concluída"
End Sub