'------------------------------------------------------------------
' Formulário: frmSimulacoes
' Descrição: Formulário para simulação de aposentadoria
'------------------------------------------------------------------

Public Sub PreencherDadosIniciais(ID_Cliente As Long)

    Dim ws As Worksheet
    Dim linha As Long
    Dim nome As String
    Dim nascimento As Date
    Dim idade As Long
    Dim tempoTotal As Double
    Dim tempoEspecial As Double
    Dim pontos As Double

    ' Buscar dados do cliente
    Set ws = Sheets("Cadastro_Clientes")
    linha = BuscarLinhaCliente(ID_Cliente)

    If linha = 0 Then
        MsgBox "Cliente não encontrado.", vbExclamation
        Exit Sub
    End If

    nome = ws.Cells(linha, 2).Value
    
    ' Valida data de nascimento
    On Error Resume Next
    nascimento = CDate(ws.Cells(linha, 7).Value)
    On Error GoTo 0
    
    ' Se a data for inválida, exibir erro e sair
    If nascimento = 0 Or Not IsDate(ws.Cells(linha, 7).Value) Then
        MsgBox "Data de nascimento inválida ou não preenchida para o cliente." & vbCrLf & _
               "Corrija o cadastro antes de fazer a simulação.", vbCritical, "Erro de Validação"
        Exit Sub
    End If
    
    ' Preencher cabeçalho
    Me.lblCliente.Caption = "Cliente: " & nome
    Me.lblID.Caption = "ID: " & ID_Cliente

    ' Calcular idade
    idade = CalcularIdade(nascimento)
    Me.txtIdadeAtual.Value = idade & " anos"

    ' Calcular tempo total de contribuição
    tempoTotal = CalcularTempo(ID_Cliente)
    Me.txtTempoTotal.Value = Format(tempoTotal, "0.00") & " anos"

    ' Calcular tempo especial convertido
    tempoEspecial = CalcularTempoEspecial(ID_Cliente)
    Me.txtTempoEspecial.Value = Format(tempoEspecial, "0.00") & " anos"

    ' Calcular pontos
    pontos = idade + tempoTotal
    Me.txtPontos.Value = Format(pontos, "0.00")

End Sub

Private Sub cmdCalcular_Click()
    ' Calcula cenário com análise automática da melhor regra
    ' Cenário rápido (15 anos) nos campos principais
    ' E adiciona cenários médio e máximo automaticamente no txtObs

    Dim ID As Long
    Dim cen1 As Collection
    Dim textoObs As String
    Dim melhorRegra As Collection
    Dim msgRegra As String

    ID = CLng(Me.Tag)
    
    ' ANÁLISE AUTOMÁTICA: Identificar a melhor regra antes de calcular
    Set melhorRegra = AnalisarMelhorRegra(ID)
    
    ' Formatar mensagem sobre a melhor regra
    Select Case melhorRegra("MelhorRegra")
        Case "TEMPO"
            msgRegra = "Aposentadoria por Tempo de Contribuição (35H/30M anos)"
        Case "IDADE"
            msgRegra = "Aposentadoria por Idade (65H/62M anos)"
        Case "PONTOS"
            msgRegra = "Regra de Pontos (105H/100M pontos)"
        Case "PEDAGIO50"
            msgRegra = "Pedágio de 50%"
        Case "PEDAGIO100"
            msgRegra = "Pedágio de 100%"
        Case Else
            msgRegra = "Regra não identificada"
    End Select
    
    ' Calcular cenário rápido (15 anos)
    Set cen1 = CalcularCenarioRapido(ID)
    
    ' Preencher campos principais com cenário 1
    If cen1("TempoFaltando") <= 0 Then
        Me.txtDireito.Value = "Sim"
        Me.txtFalta.Value = "Já elegível!"
    Else
        Me.txtDireito.Value = "Não"
        Me.txtFalta.Value = Format(cen1("TempoFaltando"), "0.00") & " anos (" & _
                           FormatarTempoExtenso(cen1("AnosFaltantes"), cen1("MesesFaltantes")) & ")"
    End If
    
    Me.txtIdadeProj.Value = cen1("IdadeAposentadoria") & " anos"
    Me.txtDataProvavel.Value = Format(cen1("DataAposentadoria"), "dd/mm/yyyy")
    Me.txtValorEstimado.Value = Format(cen1("ValorEstimado"), "R$ #,##0.00")
    
    ' Gerar texto com cenários 2 e 3 no txtObs + informação da melhor regra
    textoObs = "══════════════════════════════════════════" & vbCrLf
    textoObs = textoObs & "✓ MELHOR REGRA IDENTIFICADA:" & vbCrLf
    textoObs = textoObs & msgRegra & vbCrLf
    textoObs = textoObs & "══════════════════════════════════════════" & vbCrLf & vbCrLf
    textoObs = textoObs & GerarTextoCenariosComplementares(ID)
    
    Me.txtObs.Value = textoObs

End Sub

Private Function GerarTextoCenariosComplementares(ID_Cliente As Long) As String
    ' Gera texto formatado APENAS com cenários 2 e 3 para o txtObs
    ' O cenário 1 já está exibido nos campos principais
    
    Dim dados As Collection
    Dim cen1 As Collection
    Dim cen2 As Collection
    Dim cen3 As Collection
    Dim texto As String
    
    ' Obter dados atuais
    Set dados = ObterDadosAtuaisCliente(ID_Cliente)
    
    ' Calcular os 3 cenários
    Set cen1 = CalcularCenarioRapido(ID_Cliente)
    Set cen2 = CalcularCenarioEquilibrado(ID_Cliente)
    Set cen3 = CalcularCenarioMaximo(ID_Cliente)
    
    ' Montar texto formatado
    texto = "========== OPÇÕES DE APOSENTADORIA ===========" & vbCrLf & vbCrLf
    
    texto = texto & "SITUAÇÃO ATUAL:" & vbCrLf
    texto = texto & "• Idade: " & dados("IdadeAtual") & " anos" & vbCrLf
    texto = texto & "• Tempo Contribuído: " & Format(dados("TempoContribuido"), "0.00") & " anos" & vbCrLf
    texto = texto & vbCrLf
    
    texto = texto & "Os campos acima mostram a APOSENTADORIA RÁPIDA" & vbCrLf
    texto = texto & "(15 anos = " & cen1("Percentual") & "% do benefício = " & Format(cen1("ValorEstimado"), "R$ #,##0.00") & ")" & vbCrLf
    texto = texto & vbCrLf
    
    ' Cenário 2: Equilibrado
    texto = texto & "══════════════════════════════════════════" & vbCrLf
    texto = texto & "OPÇÃO INTERMEDIÁRIA (RECOMENDADA)" & vbCrLf
    texto = texto & "══════════════════════════════════════════" & vbCrLf
    texto = texto & "Contribuir por: " & Format(cen2("TempoNecessario"), "0.00") & " anos" & vbCrLf
    texto = texto & "Tempo faltando: " & FormatarTempoExtenso(cen2("AnosFaltantes"), cen2("MesesFaltantes")) & vbCrLf
    texto = texto & "Data prevista: " & Format(cen2("DataAposentadoria"), "dd/mm/yyyy") & vbCrLf
    texto = texto & "Idade na aposentadoria: " & cen2("IdadeAposentadoria") & " anos" & vbCrLf
    texto = texto & vbCrLf
    texto = texto & "BENEFÍCIO:" & vbCrLf
    texto = texto & "• " & cen2("Percentual") & "% da média salarial" & vbCrLf
    texto = texto & "• Valor estimado: " & Format(cen2("ValorEstimado"), "R$ #,##0.00") & vbCrLf
    
    Dim ganhoCen2 As Double
    ganhoCen2 = cen2("ValorEstimado") - cen1("ValorEstimado")
    texto = texto & "• Ganho mensal: +" & Format(ganhoCen2, "R$ #,##0.00") & vbCrLf
    texto = texto & "✓ Bom equilíbrio entre tempo e benefício" & vbCrLf
    texto = texto & vbCrLf
    
    ' Cenário 3: Máximo
    texto = texto & "══════════════════════════════════════════" & vbCrLf
    texto = texto & "BENEFÍCIO MÁXIMO (INTEGRAL)" & vbCrLf
    texto = texto & "══════════════════════════════════════════" & vbCrLf
    texto = texto & "Contribuir por: " & Format(cen3("TempoNecessario"), "0.00") & " anos" & vbCrLf
    texto = texto & "Tempo faltando: " & FormatarTempoExtenso(cen3("AnosFaltantes"), cen3("MesesFaltantes")) & vbCrLf
    texto = texto & "Data prevista: " & Format(cen3("DataAposentadoria"), "dd/mm/yyyy") & vbCrLf
    texto = texto & "Idade na aposentadoria: " & cen3("IdadeAposentadoria") & " anos" & vbCrLf
    texto = texto & vbCrLf
    texto = texto & "BENEFÍCIO:" & vbCrLf
    texto = texto & "• " & cen3("Percentual") & "% da média salarial (MÁXIMO)" & vbCrLf
    texto = texto & "• Valor estimado: " & Format(cen3("ValorEstimado"), "R$ #,##0.00") & vbCrLf
    
    Dim ganhoCen3 As Double
    ganhoCen3 = cen3("ValorEstimado") - cen1("ValorEstimado")
    texto = texto & "• Ganho mensal: +" & Format(ganhoCen3, "R$ #,##0.00") & vbCrLf
    texto = texto & "💰 Benefício integral, mas leva mais tempo" & vbCrLf
    texto = texto & vbCrLf
    
    ' Resumo comparativo
    texto = texto & "================================================" & vbCrLf
    texto = texto & "COMPARATIVO DE GANHOS:" & vbCrLf
    texto = texto & "• Rápida (15 anos): " & Format(cen1("ValorEstimado"), "R$ #,##0.00") & "/mês" & vbCrLf
    texto = texto & "• Intermediária (25 anos): " & Format(cen2("ValorEstimado"), "R$ #,##0.00") & "/mês (+" & Format((ganhoCen2 / cen1("ValorEstimado")) * 100, "0") & "%)" & vbCrLf
    texto = texto & "• Máxima (40 anos): " & Format(cen3("ValorEstimado"), "R$ #,##0.00") & "/mês (+" & Format((ganhoCen3 / cen1("ValorEstimado")) * 100, "0") & "%)" & vbCrLf
    
    GerarTextoCenariosComplementares = texto
End Function

Private Sub cmdCalcular_OLD_Click()
    ' CÓDIGO ANTIGO - MANTIDO PARA REFERÊNCIA
    ' Use cmdAnalisarTodas para análise por regras específicas

    Dim ID As Long
    Dim r As Collection
    Dim valorBeneficio As Double
    Dim tempoTotal As Double

    ID = CLng(Me.Tag)
    
    ' Obter tempo total para cálculo do benefício
    tempoTotal = CalcularTempo(ID)

    ' Verifica qual regra foi escolhida
    If Me.optTempoContribuicao.Value = True Then
        Set r = RegraTempoContribuicao(ID)
    ElseIf Me.optIdade.Value = True Then
        Set r = RegraIdade(ID)
    ElseIf Me.optPontos.Value = True Then
        Set r = RegraPontos(ID)
    ElseIf Me.optPedagio50.Value = True Then
        Set r = RegraPedagio50(ID)
    ElseIf Me.optPedagio100.Value = True Then
        Set r = RegraPedagio100(ID)
    Else
        MsgBox "Selecione uma regra.", vbExclamation
        Exit Sub
    End If

    ' Preencher resultados básicos
    Me.txtDireito.Value = r("Direito")
    Me.txtFalta.Value = Format(r("Falta"), "0.00") & " anos"
    Me.txtIdadeProj.Value = r("IdadeProjetada")
    Me.txtDataProvavel.Value = r("DataPrevista")
    Me.txtObs.Value = r("Obs")
    
    ' Calcular e exibir valor estimado do benefício
    ' Se tem direito, usa tempo atual; se não, projeta tempo futuro
    If r("Direito") = "Sim" Then
        valorBeneficio = CalcularValorBeneficio(ID, tempoTotal)
    Else
        valorBeneficio = CalcularValorBeneficio(ID, tempoTotal + r("Falta"))
    End If
    
    Me.txtValorEstimado.Value = "R$ " & Format(valorBeneficio, "#,##0.00")
    
    ' Habilitar botão de análise de cenários
    Me.cmdAnalisarCenarios.Enabled = True

End Sub

Private Sub cmdAnalisarTodas_Click()
    ' Analisa todas as regras e mostra a melhor opção
    
    Dim ID As Long
    Dim resultado As Collection
    Dim melhorRegra As String
    Dim msg As String
    
    ID = CLng(Me.Tag)
    
    Set resultado = AnalisarMelhorRegra(ID)
    melhorRegra = resultado("MelhorRegra")
    
    ' Marcar a melhor regra automaticamente
    Select Case melhorRegra
        Case "TEMPO"
            Me.optTempoContribuicao.Value = True
        Case "IDADE"
            Me.optIdade.Value = True
        Case "PONTOS"
            Me.optPontos.Value = True
        Case "PEDAGIO50"
            Me.optPedagio50.Value = True
        Case "PEDAGIO100"
            Me.optPedagio100.Value = True
    End Select
    
    ' Executar cálculo
    Call cmdCalcular_Click
    
    ' Exibir mensagem
    msg = "A melhor regra para este cliente é: " & vbCrLf & vbCrLf
    
    Select Case melhorRegra
        Case "TEMPO"
            msg = msg & "Aposentadoria por Tempo de Contribuição"
        Case "IDADE"
            msg = msg & "Aposentadoria por Idade"
        Case "PONTOS"
            msg = msg & "Regra de Pontos"
        Case "PEDAGIO50"
            msg = msg & "Pedágio de 50%"
        Case "PEDAGIO100"
            msg = msg & "Pedágio de 100%"
    End Select
    
    MsgBox msg, vbInformation, "Análise de Melhor Regra"
    
End Sub

Private Sub cmdAnalisarCenarios_Click()
    ' Botão para adicionar análise dos 3 cenários no txtObs
    ' Só é habilitado após executar cmdCalcular ou cmdAnalisarTodas
    
    Dim ID As Long
    Dim textoCompleto As String
    
    ID = CLng(Me.Tag)
    
    ' Gerar texto formatado com os 3 cenários
    textoCompleto = GerarTextoComparativoCenarios(ID)
    
    ' Substituir o conteúdo do campo de observações
    Me.txtObs.Value = textoCompleto
    
    ' Opcional: Desabilitar o botão após usar (evita cliques repetidos)
    ' Me.cmdAnalisarCenarios.Enabled = False
    
End Sub

Private Function GerarTextoComparativoCenarios(ID_Cliente As Long) As String
    ' Gera texto formatado com os 3 cenários para exibir no txtObs
    
    Dim dados As Collection
    Dim cen1 As Collection
    Dim cen2 As Collection
    Dim cen3 As Collection
    Dim texto As String
    
    ' Obter dados atuais
    Set dados = ObterDadosAtuaisCliente(ID_Cliente)
    
    ' Calcular os 3 cenários
    Set cen1 = CalcularCenarioRapido(ID_Cliente)
    Set cen2 = CalcularCenarioEquilibrado(ID_Cliente)
    Set cen3 = CalcularCenarioMaximo(ID_Cliente)
    
    ' Montar texto formatado
    texto = "========== ANÁLISE COMPARATIVA DE CENÁRIOS ==========" & vbCrLf & vbCrLf
    
    texto = texto & "DADOS ATUAIS:" & vbCrLf
    texto = texto & "• Idade: " & dados("IdadeAtual") & " anos" & vbCrLf
    texto = texto & "• Tempo Contribuído: " & Format(dados("TempoContribuido"), "0.00") & " anos" & vbCrLf
    texto = texto & vbCrLf
    
    ' Cenário 1: Rápido
    texto = texto & "══════════════════════════════════════════" & vbCrLf
    texto = texto & "CENÁRIO 1: APOSENTADORIA RÁPIDA (MÍNIMA)" & vbCrLf
    texto = texto & "══════════════════════════════════════════" & vbCrLf
    texto = texto & "Tempo necessário: " & Format(cen1("TempoNecessario"), "0.00") & " anos" & vbCrLf
    texto = texto & "Tempo faltando: " & FormatarTempoExtenso(cen1("AnosFaltantes"), cen1("MesesFaltantes")) & vbCrLf
    texto = texto & "Data prevista: " & Format(cen1("DataAposentadoria"), "dd/mm/yyyy") & vbCrLf
    texto = texto & "Idade na aposentadoria: " & cen1("IdadeAposentadoria") & " anos" & vbCrLf
    texto = texto & vbCrLf
    texto = texto & "BENEFÍCIO:" & vbCrLf
    texto = texto & "• Percentual: " & cen1("Percentual") & "% da média salarial" & vbCrLf
    texto = texto & "• Valor estimado: " & Format(cen1("ValorEstimado"), "R$ #,##0.00") & vbCrLf
    texto = texto & "⚠️  Menor valor, mas aposentadoria mais rápida" & vbCrLf
    texto = texto & vbCrLf
    
    ' Cenário 2: Equilibrado
    texto = texto & "══════════════════════════════════════════" & vbCrLf
    texto = texto & "CENÁRIO 2: APOSENTADORIA EQUILIBRADA" & vbCrLf
    texto = texto & "══════════════════════════════════════════" & vbCrLf
    texto = texto & "Tempo necessário: " & Format(cen2("TempoNecessario"), "0.00") & " anos" & vbCrLf
    texto = texto & "Tempo faltando: " & FormatarTempoExtenso(cen2("AnosFaltantes"), cen2("MesesFaltantes")) & vbCrLf
    texto = texto & "Data prevista: " & Format(cen2("DataAposentadoria"), "dd/mm/yyyy") & vbCrLf
    texto = texto & "Idade na aposentadoria: " & cen2("IdadeAposentadoria") & " anos" & vbCrLf
    texto = texto & vbCrLf
    texto = texto & "BENEFÍCIO:" & vbCrLf
    texto = texto & "• Percentual: " & cen2("Percentual") & "% da média salarial" & vbCrLf
    texto = texto & "• Valor estimado: " & Format(cen2("ValorEstimado"), "R$ #,##0.00") & vbCrLf
    texto = texto & "✓  Bom equilíbrio entre tempo e benefício (RECOMENDADO)" & vbCrLf
    texto = texto & vbCrLf
    
    ' Cenário 3: Máximo
    texto = texto & "══════════════════════════════════════════" & vbCrLf
    texto = texto & "CENÁRIO 3: APOSENTADORIA MÁXIMA (INTEGRAL)" & vbCrLf
    texto = texto & "══════════════════════════════════════════" & vbCrLf
    texto = texto & "Tempo necessário: " & Format(cen3("TempoNecessario"), "0.00") & " anos" & vbCrLf
    texto = texto & "Tempo faltando: " & FormatarTempoExtenso(cen3("AnosFaltantes"), cen3("MesesFaltantes")) & vbCrLf
    texto = texto & "Data prevista: " & Format(cen3("DataAposentadoria"), "dd/mm/yyyy") & vbCrLf
    texto = texto & "Idade na aposentadoria: " & cen3("IdadeAposentadoria") & " anos" & vbCrLf
    texto = texto & vbCrLf
    texto = texto & "BENEFÍCIO:" & vbCrLf
    texto = texto & "• Percentual: " & cen3("Percentual") & "% da média salarial" & vbCrLf
    texto = texto & "• Valor estimado: " & Format(cen3("ValorEstimado"), "R$ #,##0.00") & vbCrLf
    texto = texto & "💰 Benefício máximo, mas leva mais tempo" & vbCrLf
    texto = texto & vbCrLf
    
    ' Resumo comparativo
    texto = texto & "========================================================" & vbCrLf
    texto = texto & "COMPARATIVO RÁPIDO:" & vbCrLf
    texto = texto & "• Cenário 1 → " & Format(cen1("ValorEstimado"), "R$ #,##0.00") & " em " & FormatarTempoExtenso(cen1("AnosFaltantes"), cen1("MesesFaltantes")) & vbCrLf
    texto = texto & "• Cenário 2 → " & Format(cen2("ValorEstimado"), "R$ #,##0.00") & " em " & FormatarTempoExtenso(cen2("AnosFaltantes"), cen2("MesesFaltantes")) & vbCrLf
    texto = texto & "• Cenário 3 → " & Format(cen3("ValorEstimado"), "R$ #,##0.00") & " em " & FormatarTempoExtenso(cen3("AnosFaltantes"), cen3("MesesFaltantes")) & vbCrLf
    
    ' Calcular diferença de ganho
    Dim diferencaCen2 As Double
    Dim diferencaCen3 As Double
    diferencaCen2 = cen2("ValorEstimado") - cen1("ValorEstimado")
    diferencaCen3 = cen3("ValorEstimado") - cen1("ValorEstimado")
    
    texto = texto & vbCrLf
    texto = texto & "GANHO POR ESPERAR:" & vbCrLf
    texto = texto & "• Cenário 2 vs 1: +" & Format(diferencaCen2, "R$ #,##0.00") & " por mês" & vbCrLf
    texto = texto & "• Cenário 3 vs 1: +" & Format(diferencaCen3, "R$ #,##0.00") & " por mês" & vbCrLf
    
    GerarTextoComparativoCenarios = texto
End Function
