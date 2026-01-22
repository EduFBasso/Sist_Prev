Option Explicit

' ============================================
' SCRIPT: AtualizarSELIC
' Autor: Sistema Sist_Prev
' Data: 22/01/2026
' ============================================
' DESCRIÇÃO:
'   Baixa índices SELIC do Banco Central (série 4390)
'   e atualiza planilha "3.Indice SELIC acumulado"
'
' SÉRIE UTILIZADA:
'   4390 - Taxa SELIC acumulada no mês (% a.m.)
'   Exemplo: 1.01 significa 1.01% ao mês
'
' ⚠️ PROBLEMA COMUM:
'   Excel pode mostrar valores SEM casas decimais:
'   - API retorna: 1.01 (correto)
'   - Excel mostra: 101 (ERRADO - falta formato)
'
' ✅ SOLUÇÃO:
'   Definir .NumberFormat = "0.0000" ANTES de inserir valor
'   Isso força Excel a exibir: 1.0100
'
' REQUISITOS:
'   - Biblioteca JsonConverter instalada
'   - Conexão com internet
' ============================================

Sub AtualizarSELIC()

    Dim http As Object
    Dim json As Object
    Dim item As Object
    Dim linha As Long
    Dim url As String
    Dim ws As Worksheet
    
    ' ========================================
    ' 1. CONFIGURAÇÃO DA API
    ' ========================================
    ' Série 4390 = SELIC acumulada mensal
    ' NÃO usar série 1178 (diária - valores astronômicos)
    url = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.4390/dados?formato=json"
    
    ' ========================================
    ' 2. REQUISIÇÃO HTTP
    ' ========================================
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.send
    
    ' Debug: Mostra resposta JSON na janela Imediata
    Debug.Print http.responseText
    
    ' ========================================
    ' 3. CONVERTER JSON PARA OBJETO VBA
    ' ========================================
    Set json = JsonConverter.ParseJson(http.responseText)
    
    ' ========================================
    ' 4. PREPARAR PLANILHA
    ' ========================================
    Set ws = Sheets("3.Indice SELIC acumulado")
    
    ' ⚠️ CRÍTICO: Limpar conteúdo E FORMATO!
    ' Se só usar .ClearContents, formato % permanece
    ' e multiplica valores por 100
    ws.Range("B3:D500").ClearContents
    ws.Range("B3:D500").ClearFormats  ' Remove formato de % residual
    
    linha = 2  ' Começar na linha 2 (linha 1 é cabeçalho)
    
    ' ========================================
    ' 5. PREENCHER TABELA
    ' ========================================
    For Each item In json
        ' ----------------------------------
        ' Coluna B: Data (DD/MM/YYYY)
        ' ----------------------------------
        ws.Cells(linha, 2).Value = item("data")
        
        ' ----------------------------------
        ' Coluna C: Valor (% ao mês)
        ' ----------------------------------
        ' ⚠️ CRÍTICO: API retorna valor JÁ MULTIPLICADO POR 100!
        ' JSON vem: 1.01 MAS JsonConverter interpreta como: 101
        ' Para 1.01% ao mês, precisamos DIVIDIR por 100
        Dim valorAPI As Double
        Dim valorPercentual As Double
        
        ' Pegar valor do JSON
        valorAPI = item("valor")  ' 101 (interpretado errado)
        
        ' Dividir por 100 para obter valor real
        valorPercentual = valorAPI / 100  ' 101 / 100 = 1.01 ✅
        
        ' Debug: Descomentar para ver valores
        ' Debug.Print "API: " & valorAPI & " | Corrigido: " & valorPercentual
        
        ' Inserir valor corrigido
        ws.Cells(linha, 3).Value = valorPercentual
        
        ' Aplicar formato numérico com 4 casas
        ws.Cells(linha, 3).NumberFormat = "0.0000"
        
        ' ----------------------------------
        ' Coluna D: Fonte
        ' ----------------------------------
        ws.Cells(linha, 4).Value = "BCB API"
        
        linha = linha + 1
    Next item
    
    ' ========================================
    ' 6. REGISTRAR DATA DA ATUALIZAÇÃO
    ' ========================================
    ws.Range("G4").Value = "Atualizado em: " & Now
    
    ' ========================================
    ' 7. CONFIRMAR SUCESSO
    ' ========================================
    MsgBox "✓ Índices SELIC atualizados com sucesso!" & vbCrLf & vbCrLf & _
           "Total de registros: " & (linha - 2) & vbCrLf & vbCrLf & _
           "ATENÇÃO: Valores em % ao mês" & vbCrLf & _
           "Exemplo: 1.01 = 1.01% ao mês (não 101%!)", _
           vbInformation, "Atualização SELIC - Série 4390"

End Sub


' ============================================
' FUNÇÃO: ObterFatorSELIC
' ============================================
' DESCRIÇÃO:
'   Calcula fator de correção monetária pela SELIC
'   entre duas competências
'
' PARÂMETROS:
'   competencia - Data inicial no formato "MM/YYYY"
'                 Exemplo: "01/2020"
'   
'   competenciaBase - Data final no formato "MM/YYYY"
'                     Exemplo: "01/2026"
'
' RETORNO:
'   Fator multiplicador (Double)
'   Exemplo: 1.50 = aumenta 50%
'
' EXEMPLO DE USO:
'   fator = ObterFatorSELIC("01/2020", "01/2026")
'   valorCorrigido = 1000 * fator
'
' CÁLCULO:
'   Índice acumulado = 100 × (1 + taxa1/100) × (1 + taxa2/100) × ...
'   Fator = Índice_final / Índice_inicial
' ============================================

Function ObterFatorSELIC(competencia As String, competenciaBase As String) As Double
    
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
    Dim atingiuInicio As Boolean
    Dim atingiuFim As Boolean
    
    ' ----------------------------------
    ' 1. CONFIGURAÇÃO INICIAL
    ' ----------------------------------
    Set ws = Sheets("3.Indice SELIC acumulado")
    ultimaLinha = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    indiceInicio = 100  ' Base 100
    indiceFim = 100
    atingiuInicio = False
    atingiuFim = False
    
    ' ----------------------------------
    ' 2. ACUMULAR ÍNDICE MÊS A MÊS
    ' ----------------------------------
    For i = 2 To ultimaLinha
        ' Extrair data da célula (formato DD/MM/YYYY)
        dataAtual = ws.Cells(i, 2).Value
        
        ' Converter DD/MM/YYYY para MM/YYYY
        mesAtual = Mid(dataAtual, 4, 2)
        anoAtual = Right(dataAtual, 4)
        compAtual = mesAtual & "/" & anoAtual
        
        ' Ler taxa mensal (já está em % - exemplo: 1.01)
        valor = ws.Cells(i, 3).Value
        
        ' Acumular: Índice = Índice × (1 + taxa/100)
        ' Exemplo: 100 × (1 + 1.01/100) = 101.01
        indiceFim = indiceFim * (1 + valor / 100)
        
        ' ----------------------------------
        ' 3. MARCAR COMPETÊNCIA INICIAL
        ' ----------------------------------
        If compAtual = competencia And Not atingiuInicio Then
            indiceInicio = indiceFim
            atingiuInicio = True
        End If
        
        ' ----------------------------------
        ' 4. PARAR NA COMPETÊNCIA BASE
        ' ----------------------------------
        If compAtual = competenciaBase Then
            atingiuFim = True
            Exit For
        End If
    Next i
    
    ' ----------------------------------
    ' 5. CALCULAR FATOR FINAL
    ' ----------------------------------
    If indiceInicio > 0 And atingiuInicio Then
        ' Fator = Quanto cresceu do início até o fim
        ' Exemplo: 150 / 100 = 1.50 (cresceu 50%)
        ObterFatorSELIC = indiceFim / indiceInicio
    Else
        ' Se não encontrou competências, sem correção
        ObterFatorSELIC = 1
    End If
    
End Function


' ============================================
' SUB: ExemploCorrecaoSELIC
' ============================================
' DESCRIÇÃO:
'   Exemplo prático de como usar ObterFatorSELIC
'   para corrigir valores monetários
'
' USO:
'   Execute esta sub para ver exemplo funcionando
' ============================================

Sub ExemploCorrecaoSELIC()
    
    Dim fator As Double
    Dim valorOriginal As Double
    Dim valorCorrigido As Double
    
    ' ----------------------------------
    ' EXEMPLO: R$ 1.000 de 2020 até 2026
    ' ----------------------------------
    valorOriginal = 1000
    
    ' Calcular fator de correção
    fator = ObterFatorSELIC("01/2020", "01/2026")
    
    ' Aplicar correção
    valorCorrigido = valorOriginal * fator
    
    ' ----------------------------------
    ' MOSTRAR RESULTADO
    ' ----------------------------------
    MsgBox "📊 EXEMPLO DE CORREÇÃO MONETÁRIA PELA SELIC" & vbCrLf & vbCrLf & _
           "Valor original (Jan/2020):" & vbTab & "R$ " & Format(valorOriginal, "#,##0.00") & vbCrLf & _
           "Fator SELIC (2020→2026):" & vbTab & Format(fator, "0.0000") & "x" & vbCrLf & _
           "Valor corrigido (Jan/2026):" & vbTab & "R$ " & Format(valorCorrigido, "#,##0.00") & vbCrLf & vbCrLf & _
           "Aumento:" & vbTab & vbTab & vbTab & Format((fator - 1) * 100, "0.00") & "%", _
           vbInformation, "Exemplo Prático"
           
End Sub


' ============================================
' NOTAS IMPORTANTES
' ============================================
'
' 1. DIFERENÇA ENTRE SÉRIES BCB:
'    ┌─────────┬────────────┬─────────────────┐
'    │ Série   │ Periodo.   │ Valores 2025    │
'    ├─────────┼────────────┼─────────────────┤
'    │ 1178    │ DIÁRIA     │ 100-300 ❌      │
'    │ 4390    │ MENSAL     │ 0.8-1.3 ✅      │
'    │ 11      │ META ANUAL │ 10-13   ℹ️      │
'    └─────────┴────────────┴─────────────────┘
'
' 2. PROBLEMA COMUM DE FORMATAÇÃO:
'    API retorna: 1.01
'    Excel mostra sem .NumberFormat: 101 ❌
'    Excel mostra com .NumberFormat: 1.0100 ✅
'
' 3. INTERPRETAÇÃO DOS VALORES:
'    1.01 = 1.01% ao mês
'    Em 1 ano: (1.0101)^12 - 1 = 12.8% ao ano
'
' 4. CONTEXTO HISTÓRICO:
'    1987: 11% ao mês (hiperinflação) ✅ Normal
'    1995: 3.4% ao mês (pós-Real) ✅ Normal
'    2020: 0.4% ao mês (baixa histórica) ✅ Normal
'    2025: 1.0% ao mês (atual) ✅ Normal
'
' 5. COMPARAÇÃO INPC vs SELIC:
'    Para R$ 1.000 de Jan/1995 até Jan/2026:
'    INPC:  R$ 1.000 × 13.52 = R$ 13.520
'    SELIC: R$ 1.000 × 75.55 = R$ 75.546
'    
'    SELIC é ~5.6× maior (inclui juros, não só inflação)
'
' 6. USO CORRETO DOS ÍNDICES:
'    INPC  → Correção previdenciária (legal) ✅
'    SELIC → Juros de mora judicial ⚠️
'
' ============================================
