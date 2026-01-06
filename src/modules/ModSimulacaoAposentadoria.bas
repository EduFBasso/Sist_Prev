Attribute VB_Name = "ModSimulacaoAposentadoria"
Option Explicit

' Módulo de Simulação de Aposentadorias RGPS
' Sistema de Planejamento Previdenciário

' Estrutura para armazenar requisitos de aposentadoria
Type RequisitoAposentadoria
    tempoContribuicao As Integer
    idadeMinima As Integer
    pontos As Integer ' Para regra de pontos
    Nome As String
End Type

' Função principal de simulação
Public Function SimularAposentadorias(cpfCliente As String) As String
    Dim wsCliente As Worksheet
    Dim wsCalculo As Worksheet
    Dim wsResultados As Worksheet
    Dim linhaCliente As Long
    Dim dataNascimento As Date
    Dim sexo As String
    Dim idadeAtual As Double
    Dim tempoContribuicao As Double
    Dim resultado As String
    
    Set wsCliente = ThisWorkbook.Worksheets("Clientes")
    Set wsCalculo = ThisWorkbook.Worksheets("CalculoTempo")
    Set wsResultados = ThisWorkbook.Worksheets("ResultadosSimulacao")
    
    ' Buscar dados do cliente
    linhaCliente = ModClientes.BuscarClientePorCPF(cpfCliente)
    If linhaCliente = 0 Then
        SimularAposentadorias = "Cliente não encontrado"
        Exit Function
    End If
    
    dataNascimento = wsCliente.Cells(linhaCliente, 4).Value
    sexo = wsCliente.Cells(linhaCliente, 5).Value
    idadeAtual = CalcularIdade(dataNascimento)
    
    ' Buscar tempo de contribuição
    tempoContribuicao = BuscarTempoContribuicao(cpfCliente)
    
    ' Limpar resultados anteriores do cliente
    LimparResultadosCliente cpfCliente
    
    ' Simular cada tipo de aposentadoria
    resultado = resultado & SimularAposentadoriaPorIdade(cpfCliente, sexo, idadeAtual, tempoContribuicao) & vbCrLf
    resultado = resultado & SimularAposentadoriaPorTempoContribuicao(cpfCliente, sexo, idadeAtual, tempoContribuicao) & vbCrLf
    resultado = resultado & SimularAposentadoriaPorPontos(cpfCliente, sexo, idadeAtual, tempoContribuicao) & vbCrLf
    resultado = resultado & SimularAposentadoriaEspecial(cpfCliente, sexo, idadeAtual) & vbCrLf
    
    SimularAposentadorias = resultado
End Function

' Simular aposentadoria por idade
Private Function SimularAposentadoriaPorIdade(cpfCliente As String, sexo As String, _
                                              idadeAtual As Double, tempoContribuicao As Double) As String
    Dim idadeMinima As Integer
    Dim tempoMinimo As Integer
    Dim podeAposentar As Boolean
    Dim faltaAnos As Double
    Dim resultado As String
    
    ' Requisitos pós-reforma (2019)
    If UCase(sexo) = "M" Or UCase(sexo) = "MASCULINO" Then
        idadeMinima = 65
        tempoMinimo = 20
    Else
        idadeMinima = 62
        tempoMinimo = 15
    End If
    
    podeAposentar = (idadeAtual >= idadeMinima) And (tempoContribuicao >= tempoMinimo)
    
    If podeAposentar Then
        resultado = "✓ APOSENTADORIA POR IDADE: REQUISITOS ATENDIDOS"
        GravarResultado cpfCliente, "Aposentadoria por Idade", "Sim", 0
    Else
        If idadeAtual < idadeMinima Then
            faltaAnos = idadeMinima - idadeAtual
            resultado = "✗ Aposentadoria por Idade: Faltam " & Format(faltaAnos, "0.0") & " anos de idade"
        Else
            faltaAnos = tempoMinimo - tempoContribuicao
            resultado = "✗ Aposentadoria por Idade: Faltam " & Format(faltaAnos, "0.0") & " anos de contribuição"
        End If
        GravarResultado cpfCliente, "Aposentadoria por Idade", "Não", faltaAnos
    End If
    
    SimularAposentadoriaPorIdade = resultado
End Function

' Simular aposentadoria por tempo de contribuição
Private Function SimularAposentadoriaPorTempoContribuicao(cpfCliente As String, sexo As String, _
                                                          idadeAtual As Double, tempoContribuicao As Double) As String
    Dim tempoMinimo As Integer
    Dim idadeMinima As Integer
    Dim podeAposentar As Boolean
    Dim faltaAnos As Double
    Dim resultado As String
    
    ' Regra de transição - pedágio 100%
    If UCase(sexo) = "M" Or UCase(sexo) = "MASCULINO" Then
        tempoMinimo = 35
        idadeMinima = 60
    Else
        tempoMinimo = 30
        idadeMinima = 57
    End If
    
    podeAposentar = (tempoContribuicao >= tempoMinimo) And (idadeAtual >= idadeMinima)
    
    If podeAposentar Then
        resultado = "✓ APOSENTADORIA POR TEMPO DE CONTRIBUIÇÃO: REQUISITOS ATENDIDOS"
        GravarResultado cpfCliente, "Tempo de Contribuição", "Sim", 0
    Else
        If tempoContribuicao < tempoMinimo Then
            faltaAnos = tempoMinimo - tempoContribuicao
            resultado = "✗ Tempo de Contribuição: Faltam " & Format(faltaAnos, "0.0") & " anos de contribuição"
        Else
            faltaAnos = idadeMinima - idadeAtual
            resultado = "✗ Tempo de Contribuição: Faltam " & Format(faltaAnos, "0.0") & " anos de idade"
        End If
        GravarResultado cpfCliente, "Tempo de Contribuição", "Não", faltaAnos
    End If
    
    SimularAposentadoriaPorTempoContribuicao = resultado
End Function

' Simular aposentadoria por pontos
Private Function SimularAposentadoriaPorPontos(cpfCliente As String, sexo As String, _
                                               idadeAtual As Double, tempoContribuicao As Double) As String
    Dim pontosMinimos As Integer
    Dim tempoMinimo As Integer
    Dim pontosAtuais As Integer
    Dim podeAposentar As Boolean
    Dim faltaPontos As Integer
    Dim resultado As String
    Dim anoAtual As Integer
    
    anoAtual = Year(Date)
    
    ' Cálculo progressivo dos pontos (aumenta 1 ponto por ano)
    If UCase(sexo) = "M" Or UCase(sexo) = "MASCULINO" Then
        tempoMinimo = 35
        pontosMinimos = 100 + (anoAtual - 2019) ' 100 pontos em 2019, aumenta 1/ano até 105 em 2028
        If pontosMinimos > 105 Then pontosMinimos = 105
    Else
        tempoMinimo = 30
        pontosMinimos = 90 + (anoAtual - 2019) ' 90 pontos em 2019, aumenta 1/ano até 100 em 2033
        If pontosMinimos > 100 Then pontosMinimos = 100
    End If
    
    pontosAtuais = Int(idadeAtual + tempoContribuicao)
    podeAposentar = (pontosAtuais >= pontosMinimos) And (tempoContribuicao >= tempoMinimo)
    
    If podeAposentar Then
        resultado = "✓ APOSENTADORIA POR PONTOS: REQUISITOS ATENDIDOS (Pontos: " & pontosAtuais & ")"
        GravarResultado cpfCliente, "Aposentadoria por Pontos", "Sim", 0
    Else
        If tempoContribuicao < tempoMinimo Then
            resultado = "✗ Aposentadoria por Pontos: Tempo insuficiente"
        Else
            faltaPontos = pontosMinimos - pontosAtuais
            resultado = "✗ Aposentadoria por Pontos: Faltam " & faltaPontos & " pontos (Atual: " & pontosAtuais & " / Necessário: " & pontosMinimos & ")"
        End If
        GravarResultado cpfCliente, "Aposentadoria por Pontos", "Não", faltaPontos
    End If
    
    SimularAposentadoriaPorPontos = resultado
End Function

' Simular aposentadoria especial
Private Function SimularAposentadoriaEspecial(cpfCliente As String, sexo As String, idadeAtual As Double) As String
    Dim wsCalculo As Worksheet
    Dim tempoEspecial As Double
    Dim tempoMinimo As Integer
    Dim podeAposentar As Boolean
    Dim resultado As String
    Dim i As Long
    
    Set wsCalculo = ThisWorkbook.Worksheets("CalculoTempo")
    
    ' Buscar tempo especial
    For i = 2 To wsCalculo.Cells(wsCalculo.Rows.Count, 1).End(xlUp).Row
        If wsCalculo.Cells(i, 1).Value = cpfCliente Then
            tempoEspecial = wsCalculo.Cells(i, 5).Value
            Exit For
        End If
    Next i
    
    ' Requisitos (pode ser 15, 20 ou 25 anos conforme grau de insalubridade)
    ' Considerando grau alto (25 anos)
    tempoMinimo = 25
    
    podeAposentar = tempoEspecial >= tempoMinimo
    
    If podeAposentar Then
        resultado = "✓ APOSENTADORIA ESPECIAL: REQUISITOS ATENDIDOS (" & Format(tempoEspecial, "0.0") & " anos)"
        GravarResultado cpfCliente, "Aposentadoria Especial", "Sim", 0
    Else
        If tempoEspecial > 0 Then
            resultado = "✗ Aposentadoria Especial: Faltam " & Format(tempoMinimo - tempoEspecial, "0.0") & " anos (Tempo especial: " & Format(tempoEspecial, "0.0") & ")"
            GravarResultado cpfCliente, "Aposentadoria Especial", "Não", tempoMinimo - tempoEspecial
        Else
            resultado = "✗ Aposentadoria Especial: Nenhum tempo especial registrado"
            GravarResultado cpfCliente, "Aposentadoria Especial", "Não", tempoMinimo
        End If
    End If
    
    SimularAposentadoriaEspecial = resultado
End Function

' Função auxiliar para calcular idade
Private Function CalcularIdade(dataNascimento As Date) As Double
    CalcularIdade = (Date - dataNascimento) / 365.25
End Function

' Função auxiliar para buscar tempo de contribuição
Private Function BuscarTempoContribuicao(cpfCliente As String) As Double
    Dim wsCalculo As Worksheet
    Dim i As Long
    
    Set wsCalculo = ThisWorkbook.Worksheets("CalculoTempo")
    
    For i = 2 To wsCalculo.Cells(wsCalculo.Rows.Count, 1).End(xlUp).Row
        If wsCalculo.Cells(i, 1).Value = cpfCliente Then
            BuscarTempoContribuicao = wsCalculo.Cells(i, 2).Value
            Exit Function
        End If
    Next i
    
    BuscarTempoContribuicao = 0
End Function

' Gravar resultado da simulação
Private Sub GravarResultado(cpfCliente As String, tipoAposentadoria As String, _
                           elegivel As String, anosFaltantes As Double)
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    
    Set ws = ThisWorkbook.Worksheets("ResultadosSimulacao")
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ws.Cells(ultimaLinha, 1).Value = cpfCliente
    ws.Cells(ultimaLinha, 2).Value = tipoAposentadoria
    ws.Cells(ultimaLinha, 3).Value = elegivel
    ws.Cells(ultimaLinha, 4).Value = anosFaltantes
    ws.Cells(ultimaLinha, 5).Value = Now()
End Sub

' Limpar resultados anteriores
Private Sub LimparResultadosCliente(cpfCliente As String)
    Dim ws As Worksheet
    Dim i As Long
    
    Set ws = ThisWorkbook.Worksheets("ResultadosSimulacao")
    
    For i = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row To 2 Step -1
        If ws.Cells(i, 1).Value = cpfCliente Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub

' Inicializar planilha de resultados
Public Sub InicializarPlanilhaResultados()
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ResultadosSimulacao")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "ResultadosSimulacao"
    End If
    
    ' Criar cabeçalhos se não existirem
    If ws.Cells(1, 1).Value = "" Then
        ws.Cells(1, 1).Value = "CPF Cliente"
        ws.Cells(1, 2).Value = "Tipo Aposentadoria"
        ws.Cells(1, 3).Value = "Elegível"
        ws.Cells(1, 4).Value = "Anos Faltantes"
        ws.Cells(1, 5).Value = "Data Simulação"
        
        ' Formatar cabeçalhos
        With ws.Range("A1:E1")
            .Font.Bold = True
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With
        
        ws.Columns("A:E").AutoFit
    End If
End Sub
