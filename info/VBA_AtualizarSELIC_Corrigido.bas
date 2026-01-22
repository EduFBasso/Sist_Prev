' ============================================
' SCRIPT CORRIGIDO - Atualizar SELIC
' Série 4390 - Taxa acumulada mensal (% a.m.)
' CORREÇÃO: Formato de célula com 4 casas decimais
' ============================================

Option Explicit

Sub AtualizarSELIC()

    Dim http As Object
    Dim json As Object
    Dim item As Object
    Dim linha As Long
    Dim url As String
    Dim ws As Worksheet
    
    ' URL da API pública do Banco Central - Série 4390
    url = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.4390/dados?formato=json"
    
    ' Criar objeto HTTP
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.send
    
    Debug.Print http.responseText
    
    ' Converter JSON
    Set json = JsonConverter.ParseJson(http.responseText)
    
    ' Definir planilha
    Set ws = Sheets("3.Indice SELIC acumulado")
    
    ' Limpar tabela antiga
    ws.Range("B3:D500").ClearContents
    
    linha = 2
    
    ' Preencher tabela
    For Each item In json
        ' Data (formato DD/MM/YYYY)
        ws.Cells(linha, 2).Value = item("data")
        
        ' CORREÇÃO: Definir formato ANTES de inserir valor
        ws.Cells(linha, 3).NumberFormat = "0.0000"  ' 4 casas decimais
        
        ' Valor (% ao mês - exemplo: 1.01 significa 1.01% ao mês)
        ws.Cells(linha, 3).Value = CDbl(item("valor"))
        
        ' Fonte
        ws.Cells(linha, 4).Value = "BCB API"
        
        linha = linha + 1
    Next item
    
    ' Registrar data/hora da atualização
    ws.Range("G4").Value = "Atualizado em: " & Now
    
    ' Mensagem de sucesso com informação sobre formato
    MsgBox "Índices SELIC atualizados com sucesso!" & vbCrLf & vbCrLf & _
           "Total de registros: " & (linha - 2) & vbCrLf & _
           "Valores em % ao mês (ex: 1.01 = 1.01% ao mês)", _
           vbInformation, "Atualização SELIC"

End Sub

' ============================================
' FUNÇÃO AUXILIAR: Obter fator de correção
' ============================================
Function ObterFatorSELIC(competencia As String, competenciaBase As String) As Double
    ' Calcula fator de correção SELIC de uma competência até outra
    ' competencia: MM/YYYY (ex: "01/2020")
    ' competenciaBase: MM/YYYY (ex: "01/2026")
    '
    ' Retorna: Fator multiplicador (ex: 1.50 = aumenta 50%)
    
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim dataAtual As String
    Dim mesAtual As String
    Dim anoAtual As String
    Dim compAtual As String
    Dim indiceInicio As Double
    Dim indiceFim As Double
    Dim valor As Double
    
    Set ws = Sheets("3.Indice SELIC acumulado")
    ultimaLinha = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    ' Calcular índice acumulado desde o início
    indiceInicio = 100  ' Base 100
    indiceFim = 100
    
    Dim atingiuInicio As Boolean
    Dim atingiuFim As Boolean
    atingiuInicio = False
    atingiuFim = False
    
    ' Acumular desde o primeiro registro até competenciaBase
    For i = 2 To ultimaLinha
        dataAtual = ws.Cells(i, 2).Value  ' DD/MM/YYYY
        
        ' Extrair MM/YYYY
        mesAtual = Mid(dataAtual, 4, 2)
        anoAtual = Right(dataAtual, 4)
        compAtual = mesAtual & "/" & anoAtual
        
        ' Valor da taxa mensal (já está em %)
        valor = ws.Cells(i, 3).Value
        
        ' Acumular: Índice = Índice × (1 + taxa/100)
        indiceFim = indiceFim * (1 + valor / 100)
        
        ' Marcar quando atingir competencia inicial
        If compAtual = competencia And Not atingiuInicio Then
            indiceInicio = indiceFim
            atingiuInicio = True
        End If
        
        ' Parar quando atingir competenciaBase
        If compAtual = competenciaBase Then
            atingiuFim = True
            Exit For
        End If
    Next i
    
    ' Calcular fator
    If indiceInicio > 0 And atingiuInicio Then
        ObterFatorSELIC = indiceFim / indiceInicio
    Else
        ObterFatorSELIC = 1  ' Sem correção se não encontrou
    End If
    
End Function

' ============================================
' EXEMPLO DE USO:
' ============================================
Sub ExemploCorrecaoSELIC()
    Dim fator As Double
    Dim valorOriginal As Double
    Dim valorCorrigido As Double
    
    valorOriginal = 1000  ' R$ 1.000,00
    
    ' Corrigir de Janeiro/2020 até Janeiro/2026
    fator = ObterFatorSELIC("01/2020", "01/2026")
    valorCorrigido = valorOriginal * fator
    
    MsgBox "Valor original (01/2020): R$ " & Format(valorOriginal, "#,##0.00") & vbCrLf & _
           "Fator SELIC: " & Format(fator, "0.0000") & "x" & vbCrLf & _
           "Valor corrigido (01/2026): R$ " & Format(valorCorrigido, "#,##0.00"), _
           vbInformation, "Exemplo Correção SELIC"
End Sub
