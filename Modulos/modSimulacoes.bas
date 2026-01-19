' ============================================
' modSimulacoes - Motor de cálculo
' ============================================

Option Explicit

' Variável module-level para forçar reload do cache INPC
Private m_ForcarReloadINPC As Boolean

Function CalcularTempo(ID_Cliente As Long) As Double
    Dim ws As Worksheet
    Dim ultima As Long
    Dim i As Long
    Dim n As Long
    Dim diasTotal As Double
    Dim dtInicio As Date
    Dim dtFim As Date
    Dim k As Long
    Dim tmp As Date

    Dim datasInicio() As Date
    Dim datasFim() As Date

    Set ws = Sheets("Vinculos")

    ultima = ws.Cells(ws.Rows.count, 1).End(xlUp).Row

    For i = 2 To ultima
        If ws.Cells(i, 2).Value = ID_Cliente Then
            If IsDate(ws.Cells(i, 3).Value) Then
                dtInicio = CDate(ws.Cells(i, 3).Value)
            Else
                GoTo Proximo
            End If

            If IsDate(ws.Cells(i, 4).Value) Then
                dtFim = CDate(ws.Cells(i, 4).Value)
            Else
                dtFim = Date
            End If

            If dtFim >= dtInicio Then
                n = n + 1
                ReDim Preserve datasInicio(1 To n)
                ReDim Preserve datasFim(1 To n)
                datasInicio(n) = dtInicio
                datasFim(n) = dtFim
            End If
        End If
Proximo:
    Next i

    If n = 0 Then
        CalcularTempo = 0
        Exit Function
    End If

    Dim j As Long
    For i = 1 To n - 1
        For j = i + 1 To n
            If datasInicio(j) < datasInicio(i) Then
                tmp = datasInicio(i)
                datasInicio(i) = datasInicio(j)
                datasInicio(j) = tmp

                tmp = datasFim(i)
                datasFim(i) = datasFim(j)
                datasFim(j) = tmp
            End If
        Next j
    Next i

    Dim atualInicio As Date
    Dim atualFim As Date

    atualInicio = datasInicio(1)
    atualFim = datasFim(1)

    For k = 2 To n
        If datasInicio(k) <= atualFim Then
            If datasFim(k) > atualFim Then
                atualFim = datasFim(k)
            End If
        Else
            diasTotal = diasTotal + DateDiff("d", atualInicio, atualFim)
            atualInicio = datasInicio(k)
            atualFim = datasFim(k)
        End If
    Next k

    diasTotal = diasTotal + DateDiff("d", atualInicio, atualFim)

    CalcularTempo = diasTotal / 365.25
End Function

Function CalcularTempoAte(ID_Cliente As Long, dataLimite As Date) As Double
    ' Calcula tempo total de contribuição até uma data específica
    ' Útil para calcular tempo na data da reforma sem estimativas
    ' Considera apenas vínculos que já existiam até a data limite
    
    Dim ws As Worksheet
    Dim ultima As Long
    Dim i As Long
    Dim n As Long
    Dim diasTotal As Double
    Dim dtInicio As Date
    Dim dtFim As Date
    Dim k As Long
    Dim tmp As Date

    Dim datasInicio() As Date
    Dim datasFim() As Date

    Set ws = Sheets("Vinculos")

    ultima = ws.Cells(ws.Rows.count, 1).End(xlUp).Row

    For i = 2 To ultima
        If ws.Cells(i, 2).Value = ID_Cliente Then
            If IsDate(ws.Cells(i, 3).Value) Then
                dtInicio = CDate(ws.Cells(i, 3).Value)
            Else
                GoTo ProximoVinculo
            End If

            ' Se o vínculo começou após a data limite, ignora
            If dtInicio > dataLimite Then
                GoTo ProximoVinculo
            End If

            If IsDate(ws.Cells(i, 4).Value) Then
                dtFim = CDate(ws.Cells(i, 4).Value)
            Else
                ' Se não tem data fim, usa a menor entre hoje e dataLimite
                If Date <= dataLimite Then
                    dtFim = Date
                Else
                    dtFim = dataLimite
                End If
            End If

            ' Limita dtFim à data limite
            If dtFim > dataLimite Then
                dtFim = dataLimite
            End If

            If dtFim >= dtInicio Then
                n = n + 1
                ReDim Preserve datasInicio(1 To n)
                ReDim Preserve datasFim(1 To n)
                datasInicio(n) = dtInicio
                datasFim(n) = dtFim
            End If
        End If
ProximoVinculo:
    Next i

    If n = 0 Then
        CalcularTempoAte = 0
        Exit Function
    End If

    ' Ordenar períodos por data de início
    Dim j As Long
    For i = 1 To n - 1
        For j = i + 1 To n
            If datasInicio(j) < datasInicio(i) Then
                tmp = datasInicio(i)
                datasInicio(i) = datasInicio(j)
                datasInicio(j) = tmp

                tmp = datasFim(i)
                datasFim(i) = datasFim(j)
                datasFim(j) = tmp
            End If
        Next j
    Next i

    ' Unificar períodos sobrepostos
    Dim atualInicio As Date
    Dim atualFim As Date

    atualInicio = datasInicio(1)
    atualFim = datasFim(1)

    For k = 2 To n
        If datasInicio(k) <= atualFim Then
            If datasFim(k) > atualFim Then
                atualFim = datasFim(k)
            End If
        Else
            diasTotal = diasTotal + DateDiff("d", atualInicio, atualFim)
            atualInicio = datasInicio(k)
            atualFim = datasFim(k)
        End If
    Next k

    diasTotal = diasTotal + DateDiff("d", atualInicio, atualFim)

    CalcularTempoAte = diasTotal / 365.25
End Function

Function SimularAposentadoria(ID_Cliente As Long, regra As String) As Collection
    ' Usa GetParametro para aplicar regras
    Select Case UCase(regra)
        Case "TEMPO"
            Set SimularAposentadoria = RegraTempoContribuicao(ID_Cliente)
        Case "IDADE"
            Set SimularAposentadoria = RegraIdade(ID_Cliente)
        Case "PONTOS"
            Set SimularAposentadoria = RegraPontos(ID_Cliente)
        Case "PEDAGIO50"
            Set SimularAposentadoria = RegraPedagio50(ID_Cliente)
        Case "PEDAGIO100"
            Set SimularAposentadoria = RegraPedagio100(ID_Cliente)
        Case Else
            Set SimularAposentadoria = Nothing
    End Select
End Function

Function CalcularIdade(nascimento As Date) As Long
    CalcularIdade = DateDiff("yyyy", nascimento, Date)
    If Date < DateSerial(Year(Date), Month(nascimento), Day(nascimento)) Then
        CalcularIdade = CalcularIdade - 1
    End If
End Function

Function CalcularTempoEspecial(ID_Cliente As Long) As Double
    ' Calcula tempo especial CONVERTIDO para tempo comum
    ' Aplica fatores diferenciados por sexo (Homem/Mulher) e grau (15/20/25)
    
    Dim ws As Worksheet
    Dim wsCliente As Worksheet
    Dim ultima As Long
    Dim i As Long
    Dim linhaCliente As Long
    Dim diasTotal As Double
    Dim dtInicio As Date
    Dim dtFim As Date
    Dim grau As String
    Dim fator As Double
    Dim sexo As String

    Set ws = Sheets("Vinculos")
    Set wsCliente = Sheets("Cadastro_Clientes")
    
    ' Buscar sexo do cliente (coluna 8 em Cadastro_Clientes)
    linhaCliente = BuscarLinhaCliente(ID_Cliente)
    If linhaCliente > 0 Then
        sexo = UCase(Trim(CStr(wsCliente.Cells(linhaCliente, 8).Value)))
    Else
        sexo = "M" ' Padrão masculino se não encontrar
    End If

    ultima = ws.Cells(ws.Rows.count, 1).End(xlUp).Row

    For i = 2 To ultima
        If ws.Cells(i, 2).Value = ID_Cliente Then

            ' Coluna 6: Especial ("Sim" / "Não")
            If UCase(CStr(ws.Cells(i, 6).Value)) = "SIM" Then

                If IsDate(ws.Cells(i, 3).Value) Then
                    dtInicio = CDate(ws.Cells(i, 3).Value)
                Else
                    GoTo ProximoEspecial
                End If

                If IsDate(ws.Cells(i, 4).Value) Then
                    dtFim = CDate(ws.Cells(i, 4).Value)
                Else
                    dtFim = Date
                End If

                If dtFim >= dtInicio Then
                    grau = Trim(CStr(ws.Cells(i, 7).Value)) ' 15, 20 ou 25
                    
                    ' Buscar fator de conversão baseado em grau E sexo
                    If sexo = "M" Then
                        ' Homem: 2.33, 1.75, 1.40
                        Select Case grau
                            Case "15"
                                fator = Nz(GetParametro("Conversao_Especial_15_H"), 2.33)
                            Case "20"
                                fator = Nz(GetParametro("Conversao_Especial_20_H"), 1.75)
                            Case "25"
                                fator = Nz(GetParametro("Conversao_Especial_25_H"), 1.4)
                            Case Else
                                fator = 1
                        End Select
                    Else
                        ' Mulher: 2.00, 1.50, 1.20
                        Select Case grau
                            Case "15"
                                fator = Nz(GetParametro("Conversao_Especial_15_M"), 2)
                            Case "20"
                                fator = Nz(GetParametro("Conversao_Especial_20_M"), 1.5)
                            Case "25"
                                fator = Nz(GetParametro("Conversao_Especial_25_M"), 1.2)
                            Case Else
                                fator = 1
                        End Select
                    End If

                    If fator <= 0 Then fator = 1

                    diasTotal = diasTotal + DateDiff("d", dtInicio, dtFim) * fator
                End If
            End If
        End If
ProximoEspecial:
    Next i

    CalcularTempoEspecial = diasTotal / 365.25
End Function

Function RegraTempoContribuicao(ID_Cliente As Long) As Collection

    Dim resultado As New Collection
    Dim ws As Worksheet
    Dim linha As Long
    Dim sexo As String
    Dim tempoTotal As Double
    Dim tempoNecessario As Double
    Dim falta As Double
    Dim anosFaltantes As Long
    Dim mesesFaltantes As Long
    Dim diasFaltantes As Long
    Dim dataPrevista As Date
    Dim idadeProj As Long

    ' Buscar sexo do cliente
    Set ws = Sheets("Cadastro_Clientes")
    linha = BuscarLinhaCliente(ID_Cliente)
    sexo = ws.Cells(linha, 8).Value  ' Supondo que a coluna 6 é Sexo

    ' Tempo total já calculado
    tempoTotal = CalcularTempo(ID_Cliente)

    ' Define tempo necessário conforme sexo
    If UCase(sexo) = "M" Or UCase(sexo) = "MASCULINO" Then
        tempoNecessario = 35
    Else
        tempoNecessario = 30
    End If

    ' Calcula quanto falta
    falta = tempoNecessario - tempoTotal

    ' Se já tem direito
    If falta <= 0 Then
        resultado.Add "Sim", "Direito"
        resultado.Add tempoTotal, "TempoTotal"
        resultado.Add 0, "Falta"
        resultado.Add Date, "DataPrevista"
        resultado.Add CalcularIdade(ws.Cells(linha, 7).Value), "IdadeProjetada"
        resultado.Add "Direito adquirido pela regra antiga.", "Obs"
        Set RegraTempoContribuicao = resultado
        Exit Function
    End If

    ' Se não tem direito, calcular projeção
    ' Converte anos faltantes em dias
    Dim diasFalta As Long
    diasFalta = falta * 365.25

    dataPrevista = Date + diasFalta
    idadeProj = CalcularIdade(ws.Cells(linha, 7).Value) + falta

    resultado.Add "Não", "Direito"
    resultado.Add tempoTotal, "TempoTotal"
    resultado.Add falta, "Falta"
    resultado.Add dataPrevista, "DataPrevista"
    resultado.Add idadeProj, "IdadeProjetada"
    resultado.Add "Falta completar o tempo mínimo de contribuição.", "Obs"

    Set RegraTempoContribuicao = resultado

End Function

Function RegraIdade(ID_Cliente As Long) As Collection
    ' Aposentadoria por Idade (regra atual pós-reforma)
    Dim resultado As New Collection
    Dim ws As Worksheet
    Dim linha As Long
    Dim sexo As String
    Dim nascimento As Date
    Dim idadeAtual As Long
    Dim idadeMinima As Long
    Dim tempoTotal As Double
    Dim carenciaMinima As Long
    Dim falta As Double
    Dim dataPrevista As Date
    Dim idadeProj As Long
    
    Set ws = Sheets("Cadastro_Clientes")
    linha = BuscarLinhaCliente(ID_Cliente)
    
    sexo = ws.Cells(linha, 8).Value
    nascimento = CDate(ws.Cells(linha, 7).Value)
    idadeAtual = CalcularIdade(nascimento)
    tempoTotal = CalcularTempo(ID_Cliente)
    
    ' Buscar parâmetros
    carenciaMinima = Nz(GetParametro("Carencia_Minima")) / 12  ' Converte meses em anos
    If UCase(sexo) = "M" Or UCase(sexo) = "MASCULINO" Then
        idadeMinima = Nz(GetParametro("Idade_Minima_Homem"))
    Else
        idadeMinima = Nz(GetParametro("Idade_Minima_Mulher"))
    End If
    
    ' Verifica se já tem direito (idade mínima + carência mínima)
    If idadeAtual >= idadeMinima And tempoTotal >= carenciaMinima Then
        resultado.Add "Sim", "Direito"
        resultado.Add tempoTotal, "TempoTotal"
        resultado.Add 0, "Falta"
        resultado.Add Date, "DataPrevista"
        resultado.Add idadeAtual, "IdadeProjetada"
        resultado.Add "Direito adquirido por idade.", "Obs"
    Else
        ' Calcula o que falta
        Dim idadeFalta As Long
        Dim tempoFalta As Double
        
        idadeFalta = Application.WorksheetFunction.Max(0, idadeMinima - idadeAtual)
        tempoFalta = Application.WorksheetFunction.Max(0, carenciaMinima - tempoTotal)
        
        ' A falta é o maior dos dois
        falta = Application.WorksheetFunction.Max(idadeFalta, tempoFalta)
        
        dataPrevista = DateAdd("yyyy", falta, Date)
        idadeProj = idadeAtual + falta
        
        resultado.Add "Não", "Direito"
        resultado.Add tempoTotal, "TempoTotal"
        resultado.Add falta, "Falta"
        resultado.Add dataPrevista, "DataPrevista"
        resultado.Add idadeProj, "IdadeProjetada"
        
        If idadeFalta > 0 And tempoFalta > 0 Then
            resultado.Add "Falta idade e tempo de contribuição.", "Obs"
        ElseIf idadeFalta > 0 Then
            resultado.Add "Falta idade mínima.", "Obs"
        Else
            resultado.Add "Falta carência mínima.", "Obs"
        End If
    End If
    
    Set RegraIdade = resultado
End Function

Function RegraPontos(ID_Cliente As Long) As Collection
    ' Regra de Pontos (Idade + Tempo de Contribuição)
    Dim resultado As New Collection
    Dim ws As Worksheet
    Dim linha As Long
    Dim sexo As String
    Dim nascimento As Date
    Dim idadeAtual As Long
    Dim tempoTotal As Double
    Dim pontosAtuais As Double
    Dim pontosNecessarios As Long
    Dim tempoNecessario As Long
    Dim falta As Double
    Dim dataPrevista As Date
    Dim idadeProj As Long
    
    Set ws = Sheets("Cadastro_Clientes")
    linha = BuscarLinhaCliente(ID_Cliente)
    
    sexo = ws.Cells(linha, 8).Value
    nascimento = CDate(ws.Cells(linha, 7).Value)
    idadeAtual = CalcularIdade(nascimento)
    tempoTotal = CalcularTempo(ID_Cliente)
    
    ' Pontos atuais = idade + tempo
    pontosAtuais = idadeAtual + tempoTotal
    
    ' Buscar parâmetros
    If UCase(sexo) = "M" Or UCase(sexo) = "MASCULINO" Then
        pontosNecessarios = Nz(GetParametro("Pontos_Homem"))
        tempoNecessario = 35
    Else
        pontosNecessarios = Nz(GetParametro("Pontos_Mulher"))
        tempoNecessario = 30
    End If
    
    ' Verifica se já tem direito (pontos + tempo mínimo)
    If pontosAtuais >= pontosNecessarios And tempoTotal >= tempoNecessario Then
        resultado.Add "Sim", "Direito"
        resultado.Add tempoTotal, "TempoTotal"
        resultado.Add pontosAtuais, "Pontos"
        resultado.Add 0, "Falta"
        resultado.Add Date, "DataPrevista"
        resultado.Add idadeAtual, "IdadeProjetada"
        resultado.Add "Direito adquirido pela regra de pontos.", "Obs"
    Else
        ' Calcula o que falta
        Dim pontosFalta As Double
        Dim tempoFalta As Double
        
        pontosFalta = Application.WorksheetFunction.Max(0, pontosNecessarios - pontosAtuais)
        tempoFalta = Application.WorksheetFunction.Max(0, tempoNecessario - tempoTotal)
        
        ' A cada ano trabalhado ganha 2 pontos (1 de idade + 1 de tempo)
        ' Mas precisa atender os dois critérios
        falta = Application.WorksheetFunction.Max(pontosFalta / 2, tempoFalta)
        
        dataPrevista = DateAdd("yyyy", falta, Date)
        idadeProj = idadeAtual + falta
        
        resultado.Add "Não", "Direito"
        resultado.Add tempoTotal, "TempoTotal"
        resultado.Add pontosAtuais, "Pontos"
        resultado.Add falta, "Falta"
        resultado.Add dataPrevista, "DataPrevista"
        resultado.Add idadeProj, "IdadeProjetada"
        
        If pontosFalta > 0 And tempoFalta > 0 Then
            resultado.Add "Falta pontos e tempo de contribuição mínimo.", "Obs"
        ElseIf pontosFalta > 0 Then
            resultado.Add "Falta pontuação.", "Obs"
        Else
            resultado.Add "Falta tempo de contribuição mínimo.", "Obs"
        End If
    End If
    
    Set RegraPontos = resultado
End Function

Function RegraPedagio50(ID_Cliente As Long) As Collection
    ' Regra do Pedágio 50% (para quem estava a até 2 anos da aposentadoria na reforma)
    Dim resultado As New Collection
    Dim ws As Worksheet
    Dim linha As Long
    Dim sexo As String
    Dim nascimento As Date
    Dim idadeAtual As Long
    Dim tempoTotal As Double
    Dim tempoNecessario As Long
    Dim dataReforma As Date
    Dim tempoNaReforma As Double
    Dim tempoFaltava As Double
    Dim pedagio As Double
    Dim tempoTotalNecessario As Double
    Dim falta As Double
    Dim dataPrevista As Date
    Dim idadeProj As Long
    
    Set ws = Sheets("Cadastro_Clientes")
    linha = BuscarLinhaCliente(ID_Cliente)
    
    sexo = ws.Cells(linha, 8).Value
    nascimento = CDate(ws.Cells(linha, 7).Value)
    idadeAtual = CalcularIdade(nascimento)
    tempoTotal = CalcularTempo(ID_Cliente)
    
    ' Buscar parâmetros
    dataReforma = CDate(GetParametro("Data_Reforma"))
    If UCase(sexo) = "M" Or UCase(sexo) = "MASCULINO" Then
        tempoNecessario = 35
    Else
        tempoNecessario = 30
    End If
    
    ' Calcular tempo REAL na data da reforma (sem estimar)
    tempoNaReforma = CalcularTempoAte(ID_Cliente, dataReforma)
    
    ' Verificar elegibilidade: faltavam até 2 anos na reforma?
    tempoFaltava = tempoNecessario - tempoNaReforma
    
    If tempoFaltava > 2 Then
        resultado.Add "Não", "Direito"
        resultado.Add tempoTotal, "TempoTotal"
        resultado.Add 0, "Falta"
        resultado.Add Date, "DataPrevista"
        resultado.Add idadeAtual, "IdadeProjetada"
        resultado.Add "Não elegível: faltavam mais de 2 anos na reforma.", "Obs"
        Set RegraPedagio50 = resultado
        Exit Function
    End If
    
    ' Calcular pedágio: 50% do tempo que faltava
    pedagio = tempoFaltava * 0.5
    tempoTotalNecessario = tempoNecessario + pedagio
    
    ' Verifica se já tem direito
    If tempoTotal >= tempoTotalNecessario Then
        resultado.Add "Sim", "Direito"
        resultado.Add tempoTotal, "TempoTotal"
        resultado.Add 0, "Falta"
        resultado.Add Date, "DataPrevista"
        resultado.Add idadeAtual, "IdadeProjetada"
        resultado.Add "Direito adquirido pela regra do pedágio 50%.", "Obs"
    Else
        falta = tempoTotalNecessario - tempoTotal
        dataPrevista = DateAdd("yyyy", falta, Date)
        idadeProj = idadeAtual + falta
        
        resultado.Add "Não", "Direito"
        resultado.Add tempoTotal, "TempoTotal"
        resultado.Add falta, "Falta"
        resultado.Add dataPrevista, "DataPrevista"
        resultado.Add idadeProj, "IdadeProjetada"
        resultado.Add "Falta completar tempo + pedágio 50%.", "Obs"
    End If
    
    Set RegraPedagio50 = resultado
End Function

Function RegraPedagio100(ID_Cliente As Long) As Collection
    ' Regra do Pedágio 100% (sem idade mínima, mas pedágio de 100%)
    Dim resultado As New Collection
    Dim ws As Worksheet
    Dim linha As Long
    Dim sexo As String
    Dim nascimento As Date
    Dim idadeAtual As Long
    Dim tempoTotal As Double
    Dim tempoNecessario As Long
    Dim dataReforma As Date
    Dim tempoNaReforma As Double
    Dim tempoFaltava As Double
    Dim pedagio As Double
    Dim tempoTotalNecessario As Double
    Dim falta As Double
    Dim dataPrevista As Date
    Dim idadeProj As Long
    
    Set ws = Sheets("Cadastro_Clientes")
    linha = BuscarLinhaCliente(ID_Cliente)
    
    sexo = ws.Cells(linha, 8).Value
    nascimento = CDate(ws.Cells(linha, 7).Value)
    idadeAtual = CalcularIdade(nascimento)
    tempoTotal = CalcularTempo(ID_Cliente)
    
    ' Buscar parâmetros
    dataReforma = CDate(GetParametro("Data_Reforma"))
    If UCase(sexo) = "M" Or UCase(sexo) = "MASCULINO" Then
        tempoNecessario = 35
    Else
        tempoNecessario = 30
    End If
    
    ' Calcular tempo REAL na data da reforma (sem estimar)
    tempoNaReforma = CalcularTempoAte(ID_Cliente, dataReforma)
    
    ' Calcular tempo que faltava na reforma
    tempoFaltava = tempoNecessario - tempoNaReforma
    If tempoFaltava < 0 Then tempoFaltava = 0
    
    ' Calcular pedágio: 100% do tempo que faltava
    pedagio = tempoFaltava
    tempoTotalNecessario = tempoNecessario + pedagio
    
    ' Verifica se já tem direito
    If tempoTotal >= tempoTotalNecessario Then
        resultado.Add "Sim", "Direito"
        resultado.Add tempoTotal, "TempoTotal"
        resultado.Add 0, "Falta"
        resultado.Add Date, "DataPrevista"
        resultado.Add idadeAtual, "IdadeProjetada"
        resultado.Add "Direito adquirido pela regra do pedágio 100%.", "Obs"
    Else
        falta = tempoTotalNecessario - tempoTotal
        dataPrevista = DateAdd("yyyy", falta, Date)
        idadeProj = idadeAtual + falta
        
        resultado.Add "Não", "Direito"
        resultado.Add tempoTotal, "TempoTotal"
        resultado.Add falta, "Falta"
        resultado.Add dataPrevista, "DataPrevista"
        resultado.Add idadeProj, "IdadeProjetada"
        resultado.Add "Falta completar tempo + pedágio 100%.", "Obs"
    End If
    
    Set RegraPedagio100 = resultado
End Function

Function CalcularValorBeneficio(ID_Cliente As Long, tempoContribuicao As Double) As Double
    ' Calcula o valor estimado do benefício com base no tempo de contribuição
    ' Fórmula pós-reforma: 60% + 2% por ano acima de 15 anos (mulher) ou 20 anos (homem)
    
    Dim ws As Worksheet
    Dim linha As Long
    Dim sexo As String
    Dim coefInicial As Double
    Dim percAcrescimo As Double
    Dim tempoBase As Double
    Dim anosExcedentes As Double
    Dim coefTotal As Double
    Dim salarioMinimo As Double
    Dim tetoINSS As Double
    Dim mediaSalarialEstimada As Double
    Dim valorBeneficio As Double
    
    Set ws = Sheets("Cadastro_Clientes")
    linha = BuscarLinhaCliente(ID_Cliente)
    sexo = ws.Cells(linha, 8).Value
    
    ' Buscar parâmetros
    coefInicial = Nz(GetParametro("Coeficiente_Inicial"))
    percAcrescimo = Nz(GetParametro("Percentual_Acrescimo_Ano"))
    salarioMinimo = Nz(GetParametro("Salario_Minimo"))
    tetoINSS = Nz(GetParametro("Teto_INSS"))
    
    ' Define tempo base conforme sexo
    If UCase(sexo) = "M" Or UCase(sexo) = "MASCULINO" Then
        tempoBase = 20
    Else
        tempoBase = 15
    End If
    
    ' Calcula anos excedentes
    anosExcedentes = Application.WorksheetFunction.Max(0, tempoContribuicao - tempoBase)
    
    ' Calcula coeficiente total
    coefTotal = coefInicial + (anosExcedentes * percAcrescimo)
    If coefTotal > 100 Then coefTotal = 100
    
    ' Calcular média salarial real dos 80% maiores salários
    mediaSalarialEstimada = CalcularMediaSalarios(ID_Cliente)
    
    ' Se não houver remunerações, usar estimativa conservadora
    If mediaSalarialEstimada = 0 Then
        mediaSalarialEstimada = tetoINSS * 0.6  ' Estimativa: 60% do teto
    End If
    
    ' Calcular valor do benefício
    valorBeneficio = mediaSalarialEstimada * (coefTotal / 100)
    
    ' Aplicar limites (não pode ser menor que salário mínimo nem maior que teto)
    If valorBeneficio < salarioMinimo Then valorBeneficio = salarioMinimo
    If valorBeneficio > tetoINSS Then valorBeneficio = tetoINSS
    
    CalcularValorBeneficio = valorBeneficio
End Function

Function VerificarElegibilidadeTransicao(ID_Cliente As Long) As String
    ' Verifica se o cliente pode usar regras de transição
    ' Retorna: "TODAS", "PEDAGIO50", "PEDAGIO100", "NENHUMA"
    
    Dim ws As Worksheet
    Dim linha As Long
    Dim sexo As String
    Dim nascimento As Date
    Dim dataReforma As Date
    Dim tempoNaReforma As Double
    Dim tempoNecessario As Long
    Dim tempoFaltava As Double
    
    Set ws = Sheets("Cadastro_Clientes")
    linha = BuscarLinhaCliente(ID_Cliente)
    sexo = ws.Cells(linha, 8).Value
    nascimento = CDate(ws.Cells(linha, 7).Value)
    
    ' Buscar data da reforma
    dataReforma = CDate(GetParametro("Data_Reforma"))
    
    ' Calcular tempo de contribuição REAL na data da reforma (sem estimativas)
    tempoNaReforma = CalcularTempoAte(ID_Cliente, dataReforma)
    
    ' Define tempo necessário conforme sexo
    If UCase(sexo) = "M" Or UCase(sexo) = "MASCULINO" Then
        tempoNecessario = 35
    Else
        tempoNecessario = 30
    End If
    
    ' Calcular quanto tempo faltava na reforma
    tempoFaltava = tempoNecessario - tempoNaReforma
    
    ' Determinar elegibilidade
    If tempoFaltava <= 0 Then
        ' Já tinha direito adquirido na reforma
        VerificarElegibilidadeTransicao = "TODAS"
    ElseIf tempoFaltava <= 2 Then
        ' Elegível para ambos os pedágios
        VerificarElegibilidadeTransicao = "TODAS"
    ElseIf tempoFaltava > 2 Then
        ' Apenas pedágio 100%
        VerificarElegibilidadeTransicao = "PEDAGIO100"
    Else
        ' Não elegível para transições
        VerificarElegibilidadeTransicao = "NENHUMA"
    End If
End Function

Function AnalisarMelhorRegra(ID_Cliente As Long) As Collection
    ' Analisa todas as regras e retorna a melhor opção (menor tempo faltante)
    
    Dim resultado As New Collection
    Dim melhorRegra As String
    Dim menorFalta As Double
    Dim regras As Variant
    Dim i As Long
    Dim r As Collection
    Dim faltaAtual As Double
    Dim elegibilidade As String
    
    menorFalta = 9999
    melhorRegra = ""
    
    ' Lista de regras a verificar
    elegibilidade = VerificarElegibilidadeTransicao(ID_Cliente)
    
    ' Testar regra de tempo (direito adquirido)
    Set r = RegraTempoContribuicao(ID_Cliente)
    If r("Direito") = "Sim" Then
        resultado.Add "TEMPO", "MelhorRegra"
        resultado.Add r, "Resultado"
        Set AnalisarMelhorRegra = resultado
        Exit Function
    End If
    faltaAtual = r("Falta")
    If faltaAtual < menorFalta Then
        menorFalta = faltaAtual
        melhorRegra = "TEMPO"
        Set resultado = r
    End If
    
    ' Testar regra de idade
    Set r = RegraIdade(ID_Cliente)
    If r("Direito") = "Sim" Then
        resultado.Add "IDADE", "MelhorRegra"
        resultado.Add r, "Resultado"
        Set AnalisarMelhorRegra = resultado
        Exit Function
    End If
    faltaAtual = r("Falta")
    If faltaAtual < menorFalta Then
        menorFalta = faltaAtual
        melhorRegra = "IDADE"
        Set resultado = r
    End If
    
    ' Testar regra de pontos
    Set r = RegraPontos(ID_Cliente)
    If r("Direito") = "Sim" Then
        resultado.Add "PONTOS", "MelhorRegra"
        resultado.Add r, "Resultado"
        Set AnalisarMelhorRegra = resultado
        Exit Function
    End If
    faltaAtual = r("Falta")
    If faltaAtual < menorFalta Then
        menorFalta = faltaAtual
        melhorRegra = "PONTOS"
        Set resultado = r
    End If
    
    ' Testar pedágio 50% (se elegível)
    If elegibilidade = "TODAS" Then
        Set r = RegraPedagio50(ID_Cliente)
        If r("Direito") = "Sim" Then
            resultado.Add "PEDAGIO50", "MelhorRegra"
            resultado.Add r, "Resultado"
            Set AnalisarMelhorRegra = resultado
            Exit Function
        End If
        faltaAtual = r("Falta")
        If faltaAtual < menorFalta Then
            menorFalta = faltaAtual
            melhorRegra = "PEDAGIO50"
            Set resultado = r
        End If
    End If
    
    ' Testar pedágio 100% (se elegível)
    If elegibilidade = "TODAS" Or elegibilidade = "PEDAGIO100" Then
        Set r = RegraPedagio100(ID_Cliente)
        If r("Direito") = "Sim" Then
            resultado.Add "PEDAGIO100", "MelhorRegra"
            resultado.Add r, "Resultado"
            Set AnalisarMelhorRegra = resultado
            Exit Function
        End If
        faltaAtual = r("Falta")
        If faltaAtual < menorFalta Then
            menorFalta = faltaAtual
            melhorRegra = "PEDAGIO100"
            Set resultado = r
        End If
    End If
    
    ' Adicionar identificação da melhor regra
    resultado.Add melhorRegra, "MelhorRegra"
    
    Set AnalisarMelhorRegra = resultado
End Function

' ============================================
' FUNÇÕES PARA CÁLCULO DOS 3 CENÁRIOS
' ============================================

Function ObterDadosAtuaisCliente(ID_Cliente As Long) As Collection
    ' Retorna dados atuais do cliente
    Dim resultado As New Collection
    Dim ws As Worksheet
    Dim linha As Long
    Dim nascimento As Date
    Dim idadeAtual As Long
    Dim tempoContribuido As Double
    Dim sexo As String
    
    Set ws = Sheets("Cadastro_Clientes")
    linha = BuscarLinhaCliente(ID_Cliente)
    
    nascimento = CDate(ws.Cells(linha, 7).Value)
    sexo = ws.Cells(linha, 8).Value
    idadeAtual = CalcularIdade(nascimento)
    tempoContribuido = CalcularTempo(ID_Cliente)
    
    resultado.Add idadeAtual, "IdadeAtual"
    resultado.Add tempoContribuido, "TempoContribuido"
    resultado.Add nascimento, "DataNascimento"
    resultado.Add sexo, "Sexo"
    
    Set ObterDadosAtuaisCliente = resultado
End Function

Function CalcularCenarioRapido(ID_Cliente As Long) As Collection
    ' Cenário 1: Aposentadoria mais rápida possível (15 anos)
    Dim resultado As New Collection
    Dim dadosAtuais As Collection
    Dim tempoNecessario As Double
    Dim tempoFaltando As Double
    Dim dataAposentadoria As Date
    Dim idadeAposentadoria As Long
    Dim percentual As Double
    Dim valorEstimado As Double
    Dim anosFaltantes As Long
    Dim mesesFaltantes As Long
    
    Set dadosAtuais = ObterDadosAtuaisCliente(ID_Cliente)
    
    ' Tempo mínimo legal = 15 anos
    tempoNecessario = 15
    tempoFaltando = tempoNecessario - dadosAtuais("TempoContribuido")
    
    If tempoFaltando < 0 Then tempoFaltando = 0
    
    ' Calcular data de aposentadoria
    dataAposentadoria = DateAdd("d", tempoFaltando * 365.25, Date)
    idadeAposentadoria = dadosAtuais("IdadeAtual") + Int(tempoFaltando)
    
    ' Percentual = 60% (mínimo legal)
    percentual = 60
    
    ' Valor estimado
    valorEstimado = CalcularValorBeneficio(ID_Cliente, tempoNecessario)
    
    ' Converter tempo faltando em anos e meses
    anosFaltantes = Int(tempoFaltando)
    mesesFaltantes = Round((tempoFaltando - anosFaltantes) * 12, 0)
    
    resultado.Add tempoNecessario, "TempoNecessario"
    resultado.Add tempoFaltando, "TempoFaltando"
    resultado.Add anosFaltantes, "AnosFaltantes"
    resultado.Add mesesFaltantes, "MesesFaltantes"
    resultado.Add dataAposentadoria, "DataAposentadoria"
    resultado.Add idadeAposentadoria, "IdadeAposentadoria"
    resultado.Add percentual, "Percentual"
    resultado.Add valorEstimado, "ValorEstimado"
    
    Set CalcularCenarioRapido = resultado
End Function

Function CalcularCenarioEquilibrado(ID_Cliente As Long) As Collection
    ' Cenário 2: Balanceamento entre tempo e valor (25 anos = 78%)
    Dim resultado As New Collection
    Dim dadosAtuais As Collection
    Dim tempoNecessario As Double
    Dim tempoFaltando As Double
    Dim dataAposentadoria As Date
    Dim idadeAposentadoria As Long
    Dim percentual As Double
    Dim valorEstimado As Double
    Dim anosFaltantes As Long
    Dim mesesFaltantes As Long
    Dim sexo As String
    Dim tempoBase As Double
    
    Set dadosAtuais = ObterDadosAtuaisCliente(ID_Cliente)
    sexo = dadosAtuais("Sexo")
    
    ' Define tempo base conforme sexo
    If UCase(sexo) = "M" Or UCase(sexo) = "MASCULINO" Then
        tempoBase = 20
    Else
        tempoBase = 15
    End If
    
    ' Tempo equilibrado = 25 anos
    tempoNecessario = 25
    tempoFaltando = tempoNecessario - dadosAtuais("TempoContribuido")
    
    If tempoFaltando < 0 Then tempoFaltando = 0
    
    ' Calcular data de aposentadoria
    dataAposentadoria = DateAdd("d", tempoFaltando * 365.25, Date)
    idadeAposentadoria = dadosAtuais("IdadeAtual") + Int(tempoFaltando)
    
    ' Percentual = 60% + 2% × (25 - tempoBase)
    percentual = 60 + (2 * (tempoNecessario - tempoBase))
    
    ' Valor estimado
    valorEstimado = CalcularValorBeneficio(ID_Cliente, tempoNecessario)
    
    ' Converter tempo faltando em anos e meses
    anosFaltantes = Int(tempoFaltando)
    mesesFaltantes = Round((tempoFaltando - anosFaltantes) * 12, 0)
    
    resultado.Add tempoNecessario, "TempoNecessario"
    resultado.Add tempoFaltando, "TempoFaltando"
    resultado.Add anosFaltantes, "AnosFaltantes"
    resultado.Add mesesFaltantes, "MesesFaltantes"
    resultado.Add dataAposentadoria, "DataAposentadoria"
    resultado.Add idadeAposentadoria, "IdadeAposentadoria"
    resultado.Add percentual, "Percentual"
    resultado.Add valorEstimado, "ValorEstimado"
    
    Set CalcularCenarioEquilibrado = resultado
End Function

Function CalcularCenarioMaximo(ID_Cliente As Long) As Collection
    ' Cenário 3: Benefício máximo (40 anos = 100%)
    Dim resultado As New Collection
    Dim dadosAtuais As Collection
    Dim tempoNecessario As Double
    Dim tempoFaltando As Double
    Dim dataAposentadoria As Date
    Dim idadeAposentadoria As Long
    Dim percentual As Double
    Dim valorEstimado As Double
    Dim anosFaltantes As Long
    Dim mesesFaltantes As Long
    
    Set dadosAtuais = ObterDadosAtuaisCliente(ID_Cliente)
    
    ' Tempo máximo = 40 anos
    tempoNecessario = 40
    tempoFaltando = tempoNecessario - dadosAtuais("TempoContribuido")
    
    If tempoFaltando < 0 Then tempoFaltando = 0
    
    ' Calcular data de aposentadoria
    dataAposentadoria = DateAdd("d", tempoFaltando * 365.25, Date)
    idadeAposentadoria = dadosAtuais("IdadeAtual") + Int(tempoFaltando)
    
    ' Percentual = 100% (máximo)
    percentual = 100
    
    ' Valor estimado
    valorEstimado = CalcularValorBeneficio(ID_Cliente, tempoNecessario)
    
    ' Converter tempo faltando em anos e meses
    anosFaltantes = Int(tempoFaltando)
    mesesFaltantes = Round((tempoFaltando - anosFaltantes) * 12, 0)
    
    resultado.Add tempoNecessario, "TempoNecessario"
    resultado.Add tempoFaltando, "TempoFaltando"
    resultado.Add anosFaltantes, "AnosFaltantes"
    resultado.Add mesesFaltantes, "MesesFaltantes"
    resultado.Add dataAposentadoria, "DataAposentadoria"
    resultado.Add idadeAposentadoria, "IdadeAposentadoria"
    resultado.Add percentual, "Percentual"
    resultado.Add valorEstimado, "ValorEstimado"
    
    Set CalcularCenarioMaximo = resultado
End Function

Function FormatarTempoExtenso(anos As Long, meses As Long) As String
    ' Formata tempo em texto legível
    Dim resultado As String
    
    If anos = 0 And meses = 0 Then
        resultado = "Já elegível!"
    ElseIf anos = 0 Then
        resultado = meses & " " & IIf(meses = 1, "mês", "meses")
    ElseIf meses = 0 Then
        resultado = anos & " " & IIf(anos = 1, "ano", "anos")
    Else
        resultado = anos & " " & IIf(anos = 1, "ano", "anos") & " e " & meses & " " & IIf(meses = 1, "mês", "meses")
    End If
    
    FormatarTempoExtenso = resultado
End Function

Function CalcularMediaSalarios(ID_Cliente As Long) As Double
    ' Calcula a média dos 80% maiores salários do cliente
    ' baseado nas remunerações importadas do CNIS
    ' COM APLICACAO DE CORRECAO INPC
    
    Dim ws As Worksheet
    Dim ultima As Long
    Dim i As Long
    Dim salarios() As Double
    Dim n As Long
    Dim total As Double
    Dim quantidadeConsiderar As Long
    
    On Error Resume Next
    Set ws = Sheets("Remuneracoes")
    On Error GoTo 0
    
    If ws Is Nothing Then
        CalcularMediaSalarios = 0
        Exit Function
    End If
    
    ' Coletar todos os salários do cliente COM CORRECAO INPC
    ultima = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    n = 0
    
    For i = 2 To ultima
        If ws.Cells(i, 3).Value = ID_Cliente Then
            Dim valor As Double
            Dim competencia As String
            Dim valorCorrigido As Double
            
            valor = ws.Cells(i, 5).Value          ' Coluna E = Valor
            competencia = ws.Cells(i, 4).Value    ' Coluna D = Competencia (MM/YYYY)
            
            If valor > 0 Then
                ' Aplicar correção monetária ao valor (índice configurável)
                valorCorrigido = AplicarCorrecaoMonetaria(valor, competencia)
                
                n = n + 1
                ReDim Preserve salarios(1 To n)
                salarios(n) = valorCorrigido
            End If
        End If
    Next i
    
    If n = 0 Then
        CalcularMediaSalarios = 0
        Exit Function
    End If
    
    ' Ordenar salários em ordem decrescente (bubble sort simples)
    Dim j As Long
    Dim temp As Double
    
    For i = 1 To n - 1
        For j = i + 1 To n
            If salarios(j) > salarios(i) Then
                temp = salarios(i)
                salarios(i) = salarios(j)
                salarios(j) = temp
            End If
        Next j
    Next i
    
    ' Calcular média dos 80% maiores
    quantidadeConsiderar = Application.WorksheetFunction.Max(1, Int(n * 0.8))
    
    total = 0
    For i = 1 To quantidadeConsiderar
        total = total + salarios(i)
    Next i
    
    CalcularMediaSalarios = total / quantidadeConsiderar
    
End Function

' ==============================================================================
' FUNCAO: AplicarCorrecaoMonetaria
' DESCRICAO: Ponto único para escolher o índice de correção monetária.
'            Hoje: INPC (implementado). Futuro: SELIC (a implementar).
' PARAMETROS:
'   - valor: valor original
'   - competencia: data no formato MM/YYYY
' RETORNO: Valor corrigido conforme índice selecionado
' ==============================================================================
Public Function AplicarCorrecaoMonetaria(ByVal valor As Double, ByVal competencia As String) As Double
    Dim indice As String
    indice = UCase(Trim(CStr(GetParametro("Indice_Correcao_Remuneracoes"))))

    If indice = "" Or indice = "0" Then
        indice = "INPC" ' Padrão seguro
    End If

    Select Case indice
        Case "INPC"
            AplicarCorrecaoMonetaria = AplicarCorrecaoINPC(valor, competencia)
        Case "SELIC"
            AplicarCorrecaoMonetaria = AplicarCorrecaoSELIC(valor, competencia)
        Case Else
            ' Índice desconhecido: não corrige (evita quebrar simulação)
            AplicarCorrecaoMonetaria = valor
    End Select
End Function

' ==============================================================================
' FUNCAO: AplicarCorrecaoSELIC
' DESCRICAO: Placeholder para correção por SELIC (não implementado neste projeto).
'            Mantém o motor extensível sem travar o uso do sistema.
' ==============================================================================
Private Function AplicarCorrecaoSELIC(ByVal valor As Double, ByVal competencia As String) As Double
    Static avisoMostrado As Boolean

    If Not avisoMostrado Then
        avisoMostrado = True
        MsgBox "Correção por SELIC ainda não está implementada." & vbCrLf & vbCrLf & _
               "Parâmetro 'Indice_Correcao_Remuneracoes' está como SELIC, mas o sistema irá manter os valores sem correção (por enquanto)." & vbCrLf & _
               "Para usar INPC, defina o parâmetro como INPC em Config_Regras.", _
               vbExclamation, "Aviso - SELIC pendente"
    End If

    AplicarCorrecaoSELIC = valor
End Function

' ================================================================================
' FUNCAO: AtualizarIndicesINPC
' DESCRICAO: Executa script Python para baixar indices INPC do BCB e atualizar
'            fatores de correcao monetaria
' RETORNO: True se sucesso, False se erro
' ================================================================================
Public Function AtualizarIndicesINPC() As Boolean
    On Error GoTo ErroHandler
    
    Dim caminhoExcel As String
    Dim caminhoExe As String
    Dim caminhoScript As String
    Dim comando As String
    Dim resultado As Long
    
    ' Obter caminho do arquivo Excel
    caminhoExcel = ThisWorkbook.Path
    
    ' Detectar modo: PRODUCAO (.exe) ou DESENVOLVIMENTO (.py)
    caminhoExe = caminhoExcel & Application.PathSeparator & "bin" & _
                 Application.PathSeparator & "atualizar_inpc.exe"
    caminhoScript = caminhoExcel & Application.PathSeparator & "atualizar_inpc.py"
    
    If Dir(caminhoExe) <> "" Then
        ' MODO PRODUCAO: Executavel standalone (distribuicao ao cliente)
        comando = """" & caminhoExe & """"
    ElseIf Dir(caminhoScript) <> "" Then
        ' MODO DESENVOLVIMENTO: Script Python (requer Python instalado)
        comando = "python """ & caminhoScript & """"
    Else
        MsgBox "Erro: Nenhum executavel encontrado!" & vbCrLf & vbCrLf & _
               "Procurado em:" & vbCrLf & _
               "1) " & caminhoExe & vbCrLf & _
               "2) " & caminhoScript, vbCritical, "Erro"
        AtualizarIndicesINPC = False
        Exit Function
    End If
    
    ' Informar usuario
    Application.ScreenUpdating = False
    Application.StatusBar = "Conectando ao Banco Central do Brasil..."
    
    ' Executar script Python ou .exe
    resultado = Shell(comando, vbNormalFocus)
    
    ' Aguardar alguns segundos para conclusao
    Application.Wait (Now + TimeValue("00:00:05"))
    
    ' Verificar se arquivo foi criado
    Dim arquivoSaida As String
    arquivoSaida = caminhoExcel & Application.PathSeparator & "saida" & _
                   Application.PathSeparator & "inpc_fatores.csv"
    
    If Dir(arquivoSaida) = "" Then
        MsgBox "Erro: Arquivo inpc_fatores.csv nao foi gerado!" & vbCrLf & _
               "Verifique a execucao do script Python.", vbCritical, "Erro"
        AtualizarIndicesINPC = False
        Application.StatusBar = False
        Application.ScreenUpdating = True
        Exit Function
    End If
    
    ' Atualizar parametro no Config_Regras
    Call SetParametro("Data_Atualizacao_INPC", CStr(Date), _
                      "Data da ultima atualizacao dos indices INPC do BCB")
    
    ' IMPORTANTE: Limpar cache do AplicarCorrecaoINPC para forcar reload do CSV
    Call LimparCacheINPC
    
    ' Restaurar interface
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    ' Sucesso
    MsgBox "Indices INPC atualizados com sucesso!" & vbCrLf & _
           "Data: " & Format(Date, "dd/mm/yyyy") & vbCrLf & vbCrLf & _
           "Os calculos de simulacao agora usarao valores corrigidos.", _
           vbInformation, "Atualizacao Concluida"
    
    AtualizarIndicesINPC = True
    Exit Function
    
ErroHandler:
    MsgBox "Erro ao atualizar indices INPC:" & vbCrLf & _
           "Numero: " & Err.Number & vbCrLf & _
           "Descricao: " & Err.Description, vbCritical, "Erro"
    Application.StatusBar = False
    Application.ScreenUpdating = True
    AtualizarIndicesINPC = False
End Function

' ================================================================================
' FUNCAO: AplicarCorrecaoINPC
' DESCRICAO: Aplica correcao monetaria INPC a um valor baseado na competencia
' PARAMETROS:
'   - valor: valor original a ser corrigido
'   - competencia: data no formato MM/YYYY
' RETORNO: Valor corrigido ou valor original se fator nao encontrado
' ================================================================================
Private Function AplicarCorrecaoINPC(ByVal valor As Double, ByVal competencia As String) As Double
    On Error Resume Next
    
    Static fatoresINPC As Object ' Dictionary para cache
    Static ultimaLeitura As Date
    
    ' Inicializar dictionary se necessario, passou mais de 1 dia, ou foi solicitado reload
    If fatoresINPC Is Nothing Or DateDiff("d", ultimaLeitura, Date) > 0 Or m_ForcarReloadINPC Then
        Set fatoresINPC = CreateObject("Scripting.Dictionary")
        m_ForcarReloadINPC = False ' Resetar flag
        
        Dim caminhoArquivo As String
        Dim linha As String
        Dim partes() As String
        Dim fileNum As Integer
        
        caminhoArquivo = ThisWorkbook.Path & Application.PathSeparator & "saida" & _
                        Application.PathSeparator & "inpc_fatores.csv"
        
        ' Verificar se arquivo existe
        If Dir(caminhoArquivo) = "" Then
            ' Arquivo nao existe, retornar valor sem correcao
            AplicarCorrecaoINPC = valor
            Exit Function
        End If
        
        ' Ler arquivo CSV
        fileNum = FreeFile
        Open caminhoArquivo For Input As #fileNum
        
        ' Pular cabecalho
        Line Input #fileNum, linha
        
        Dim contadorLinhas As Integer
        contadorLinhas = 0
        
        ' Ler fatores
        Do While Not EOF(fileNum)
            Line Input #fileNum, linha
            
            If Len(Trim(linha)) > 0 Then
                partes = Split(linha, ";")
                If UBound(partes) >= 1 Then
                    ' Adicionar ao dictionary: chave=competencia, valor=fator
                    fatoresINPC(Trim(partes(0))) = CDbl(Replace(partes(1), ".", ","))
                    contadorLinhas = contadorLinhas + 1
                End If
            End If
        Loop
        
        Close #fileNum
        ultimaLeitura = Date
    End If
    
    ' Aplicar fator se existir
    If fatoresINPC.Exists(competencia) Then
        AplicarCorrecaoINPC = valor * fatoresINPC(competencia)
    Else
        ' Fator nao encontrado, retornar valor original
        AplicarCorrecaoINPC = valor
    End If
    
End Function

' ================================================================================
' FUNCAO: LimparCacheINPC
' DESCRICAO: Limpa cache de fatores INPC forçando reload do CSV na próxima consulta
' ================================================================================
Public Sub LimparCacheINPC()
    ' Define flag para forçar reload na próxima chamada de AplicarCorrecaoINPC
    m_ForcarReloadINPC = True
End Sub
