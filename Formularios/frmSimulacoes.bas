'------------------------------------------------------------------
' Formulário: frmSimulacoes
' Descrição: Formulário para simulação de aposentadoria
'------------------------------------------------------------------

Private Sub UserForm_Initialize()
    ' Carregar estado do modo automático ao abrir o formulário
    Call AtualizarBotaoToggle
End Sub

Private Sub UserForm_Activate()
    ' Garantir que o botão esteja atualizado quando formulário ganhar foco
    Call AtualizarBotaoToggle
    Call AtualizarEstadoBotoes
End Sub

Private Sub AtualizarBotaoToggle()
    ' Atualiza visual do botão toggle baseado no parâmetro salvo
    Dim modoAuto As Variant
    modoAuto = GetParametro("ModoAutoSimulacao")
    
    ' Se parâmetro não existe (retorna 0) ou está vazio, criar com valor padrão TRUE
    If modoAuto = 0 Or modoAuto = "" Or IsNull(modoAuto) Then
        Call SetParametro("ModoAutoSimulacao", "TRUE", "Modo automático nas simulações (TRUE/FALSE)")
        modoAuto = "TRUE"
    End If
    
    ' Converter para string e limpar
    Dim modoStr As String
    modoStr = UCase(Trim(CStr(modoAuto)))
    
    ' Atualizar Caption e cor do botão baseado no estado
    ' Aceitar TRUE (inglês) ou VERDADEIRO (português)
    If modoStr = "TRUE" Or modoStr = "VERDADEIRO" Then
        Me.cmdToggleAuto.Caption = "Auto"
        Me.cmdToggleAuto.BackColor = RGB(144, 238, 144)  ' Verde padrão
    Else
        Me.cmdToggleAuto.Caption = "Manual"
        Me.cmdToggleAuto.BackColor = RGB(255, 218, 185)  ' Laranja claro
    End If
End Sub

Private Sub AtualizarEstadoBotoes()
    ' Habilita/desabilita OptionButtons baseado no modo automático
    ' Modo AUTO: botões desabilitados (sistema escolhe automaticamente)
    ' Modo MANUAL: botões habilitados (usuário escolhe)
    
    Dim modoAuto As Variant
    Dim habilitar As Boolean
    
    modoAuto = GetParametro("ModoAutoSimulacao")
    
    ' Se não existe, assumir TRUE (modo AUTO)
    If modoAuto = 0 Or modoAuto = "" Or IsNull(modoAuto) Then
        modoAuto = "TRUE"
    End If
    
    ' Se modo AUTO, desabilitar botões (sistema decide)
    ' Se modo MANUAL, habilitar botões (usuário decide)
    Dim modoStr As String
    modoStr = UCase(Trim(CStr(modoAuto)))
    habilitar = (modoStr <> "TRUE" And modoStr <> "VERDADEIRO")
    
    ' Atualizar estado de todos os OptionButtons
    Me.optTempoContribuicao.Enabled = habilitar
    Me.optIdade.Enabled = habilitar
    Me.optPontos.Enabled = habilitar
    Me.optPedagio50.Enabled = habilitar
    Me.optPedagio100.Enabled = habilitar
    Me.optEspecial.Enabled = habilitar
End Sub

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
    
    ' PRÉ-SELECIONAR o optionButton correspondente à melhor regra
    Select Case melhorRegra("MelhorRegra")
        Case "TEMPO"
            Me.optTempoContribuicao.Value = True
            msgRegra = "Aposentadoria por Tempo de Contribuição (35H/30M anos)"
        Case "IDADE"
            Me.optIdade.Value = True
            msgRegra = "Aposentadoria por Idade (65H/62M anos)"
        Case "PONTOS"
            Me.optPontos.Value = True
            msgRegra = "Regra de Pontos (105H/100M pontos)"
        Case "PEDAGIO50"
            Me.optPedagio50.Value = True
            msgRegra = "Pedágio de 50%"
        Case "PEDAGIO100"
            Me.optPedagio100.Value = True
            msgRegra = "Pedágio de 100%"
        Case "ESPECIAL"
            Me.optEspecial.Value = True
            msgRegra = "Aposentadoria Especial (15/20/25 anos)"
            MsgBox "AVISO: A regra 'Aposentadoria Especial' foi identificada, mas ainda não está completamente implementada." & vbCrLf & vbCrLf & _
                   "Consulte o escritório para validar se esta regra se aplica ao caso.", _
                   vbExclamation, "Implementação Pendente"
        Case Else
            msgRegra = "Regra não identificada"
    End Select
    
    ' ACESSIBILIDADE: No modo AUTO, desabilitar todos exceto o selecionado
    ' Deixa mais visível qual regra foi escolhida pelo sistema
    Dim modoAuto As Variant
    modoAuto = GetParametro("ModoAutoSimulacao")
    If modoAuto = 0 Or modoAuto = "" Or IsNull(modoAuto) Then modoAuto = "TRUE"
    
    Dim modoStr As String
    modoStr = UCase(Trim(CStr(modoAuto)))
    
    If modoStr = "TRUE" Or modoStr = "VERDADEIRO" Then
        ' Desabilitar todos primeiro
        Me.optTempoContribuicao.Enabled = False
        Me.optIdade.Enabled = False
        Me.optPontos.Enabled = False
        Me.optPedagio50.Enabled = False
        Me.optPedagio100.Enabled = False
        Me.optEspecial.Enabled = False
        
        ' Habilitar APENAS o selecionado (fica mais visível)
        Select Case melhorRegra("MelhorRegra")
            Case "TEMPO": Me.optTempoContribuicao.Enabled = True
            Case "IDADE": Me.optIdade.Enabled = True
            Case "PONTOS": Me.optPontos.Enabled = True
            Case "PEDAGIO50": Me.optPedagio50.Enabled = True
            Case "PEDAGIO100": Me.optPedagio100.Enabled = True
            Case "ESPECIAL": Me.optEspecial.Enabled = True
        End Select
    End If
    
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
    textoObs = "MELHOR REGRA IDENTIFICADA:" & vbCrLf
    textoObs = textoObs & msgRegra & vbCrLf & vbCrLf
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
    texto = "OPÇÕES DE APOSENTADORIA" & vbCrLf & vbCrLf
    
    texto = texto & "SITUAÇÃO ATUAL:" & vbCrLf
    texto = texto & "• Idade: " & dados("IdadeAtual") & " anos" & vbCrLf
    texto = texto & "• Tempo Contribuído: " & Format(dados("TempoContribuido"), "0.00") & " anos" & vbCrLf
    texto = texto & vbCrLf
    
    ' Cenário 1: Rápido (DETALHADO)
    texto = texto & "APOSENTADORIA RÁPIDA" & vbCrLf
    texto = texto & "Contribuir por: " & Format(cen1("TempoNecessario"), "0.00") & " anos" & vbCrLf
    texto = texto & "Tempo faltando: " & FormatarTempoExtenso(cen1("AnosFaltantes"), cen1("MesesFaltantes")) & vbCrLf
    texto = texto & "Data prevista: " & Format(cen1("DataAposentadoria"), "dd/mm/yyyy") & vbCrLf
    texto = texto & "Idade na aposentadoria: " & cen1("IdadeAposentadoria") & " anos" & vbCrLf
    texto = texto & vbCrLf
    texto = texto & "BENEFÍCIO:" & vbCrLf
    texto = texto & "• " & cen1("Percentual") & "% da média salarial" & vbCrLf
    texto = texto & "• Valor estimado: " & Format(cen1("ValorEstimado"), "R$ #,##0.00") & vbCrLf
    texto = texto & "Menor tempo de espera, benefício proporcional" & vbCrLf
    texto = texto & vbCrLf
    
    ' Cenário 2: Equilibrado
    texto = texto & "APOSENTADORIA SEMI-INTEGRAL (RECOMENDADA)" & vbCrLf
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
    texto = texto & "Bom equilíbrio entre tempo e benefício" & vbCrLf
    texto = texto & vbCrLf
    
    ' Cenário 3: Máximo
    texto = texto & "APOSENTADORIA INTEGRAL" & vbCrLf
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
    texto = texto & "Benefício integral, mas leva mais tempo" & vbCrLf
    texto = texto & vbCrLf
    
    ' Resumo comparativo
    texto = texto & "COMPARATIVO DE GANHOS:" & vbCrLf
    texto = texto & "• Rápida (15 anos): " & Format(cen1("ValorEstimado"), "R$ #,##0.00") & "/mês" & vbCrLf
    texto = texto & "• Intermediária (25 anos): " & Format(cen2("ValorEstimado"), "R$ #,##0.00") & "/mês (+" & Format((ganhoCen2 / cen1("ValorEstimado")) * 100, "0") & "%)" & vbCrLf
    texto = texto & "• Máxima (40 anos): " & Format(cen3("ValorEstimado"), "R$ #,##0.00") & "/mês (+" & Format((ganhoCen3 / cen1("ValorEstimado")) * 100, "0") & "%)" & vbCrLf
    texto = texto & vbCrLf
    texto = texto & String(57, "=") & vbCrLf
    texto = texto & "OBSERVAÇÕES DO ADVOGADO:" & vbCrLf
    texto = texto & "(Espaço para anotações adicionais)" & vbCrLf
    texto = texto & vbCrLf
    
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

Private Sub cmdToggleAuto_Click()
    ' Alterna entre modo automático e manual
    ' Modo automático: cmdCalcular executa análise completa automaticamente
    ' Modo manual: usuário precisa escolher regra manualmente
    
    Dim modoAtual As Variant
    Dim novoModo As String
    
    modoAtual = GetParametro("ModoAutoSimulacao")
    
    ' Se parâmetro não existe (retorna 0), assumir padrão TRUE
    If modoAtual = 0 Or modoAtual = "" Or IsNull(modoAtual) Then
        modoAtual = "TRUE"
    End If
    
    ' Alternar estado (aceitar TRUE ou VERDADEIRO)
    Dim modoStr As String
    modoStr = UCase(Trim(CStr(modoAtual)))
    
    If modoStr = "TRUE" Or modoStr = "VERDADEIRO" Then
        novoModo = "FALSE"
    Else
        novoModo = "TRUE"
    End If
    
    ' IMPORTANTE: Salvar ANTES do MsgBox para evitar conflito com UserForm_Activate
    Call SetParametro("ModoAutoSimulacao", novoModo, "Modo automático nas simulações (TRUE/FALSE)")
    
    ' Atualizar visual manualmente (não chamar AtualizarBotaoToggle para evitar releitura)
    If novoModo = "TRUE" Then
        Me.cmdToggleAuto.Caption = "Auto"
        Me.cmdToggleAuto.BackColor = RGB(144, 238, 144)  ' Verde padrão
    Else
        Me.cmdToggleAuto.Caption = "Manual"
        Me.cmdToggleAuto.BackColor = RGB(255, 218, 185)  ' Laranja claro
    End If
    
    ' Atualizar estado dos OptionButtons
    Call AtualizarEstadoBotoes
    
    ' Mensagem DEPOIS de salvar (evita conflito de foco)
    If novoModo = "FALSE" Then
        MsgBox "Modo MANUAL ativado!" & vbCrLf & vbCrLf & _
               "Agora você precisa selecionar a regra manualmente antes de calcular.", _
               vbInformation, "Modo Alterado"
    Else
        MsgBox "Modo AUTOMÁTICO ativado!" & vbCrLf & vbCrLf & _
               "O sistema identificará automaticamente a melhor regra ao calcular.", _
               vbInformation, "Modo Alterado"
    End If
    
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

' ============================================
' Botão Imprimir Simulação
' ============================================
Private Sub cmdImprimir_Click()
    ' Imprime o conteúdo da simulação com cabeçalho da empresa (se configurado)
    
    If Trim(Me.txtObs.Value) = "" Then
        MsgBox "Não há simulação para imprimir!" & vbCrLf & vbCrLf & _
               "Execute uma simulação primeiro usando o botão 'Calcular'.", _
               vbExclamation, "Nenhum Conteúdo"
        Exit Sub
    End If
    
    ' Gerar relatório formatado
    Call ImprimirSimulacao(Me.lblCliente.Caption, Me.lblID.Caption, Me.txtObs.Value)
End Sub

' ============================================
' Função para imprimir simulação
' ============================================
Private Sub ImprimirSimulacao(nomeCliente As String, idCliente As String, conteudo As String)
    ' Cria um documento Word temporário e abre a visualização de impressão
    
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim cabecalho As String
    Dim rodape As String
    Dim textoCompleto As String
    Dim rng As Object
    
    On Error GoTo ErroWord
    
    ' Gerar texto completo (cabeçalho + conteúdo + rodapé)
    cabecalho = GerarCabecalhoImpressao(nomeCliente, idCliente)
    rodape = GerarRodapeImpressao()
    
    textoCompleto = ""
    If cabecalho <> "" Then
        textoCompleto = cabecalho & vbCrLf
    End If
    textoCompleto = textoCompleto & conteudo & vbCrLf
    If rodape <> "" Then
        textoCompleto = textoCompleto & rodape
    End If
    
    ' Criar instância do Word
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True
    
    ' Criar novo documento
    Set wordDoc = wordApp.Documents.Add
    
    ' Configurar página A4 (sem usar constantes)
    On Error Resume Next
    wordDoc.PageSetup.PaperSize = 9
    wordDoc.PageSetup.Orientation = 0
    wordDoc.PageSetup.TopMargin = 36
    wordDoc.PageSetup.BottomMargin = 36
    wordDoc.PageSetup.LeftMargin = 36
    wordDoc.PageSetup.RightMargin = 36
    On Error GoTo ErroWord
    
    ' Inserir texto no documento
    wordDoc.Content.Text = textoCompleto
    
    ' Formatar todo o documento com Courier New (base)
    On Error Resume Next
    wordDoc.Content.Font.Name = "Courier New"
    wordDoc.Content.Font.Size = 10
    On Error GoTo ErroWord
    
    ' ========== APLICAR FORMATAÇÃO ESPECIAL ==========
    On Error Resume Next  ' Continuar mesmo se formatação falhar
    
    ' Título principal: SIMULAÇÃO DE APOSENTADORIA
    Set rng = wordDoc.Content
    rng.Find.ClearFormatting
    rng.Find.Text = "SIMULAÇÃO DE APOSENTADORIA"
    rng.Find.Forward = True
    If rng.Find.Execute Then
        rng.Font.Size = 14
        rng.Font.Bold = True
        rng.ParagraphFormat.Alignment = 1
    End If
    
    ' Subtítulo: MELHOR REGRA IDENTIFICADA:
    Set rng = wordDoc.Content
    rng.Find.ClearFormatting
    rng.Find.Text = "MELHOR REGRA IDENTIFICADA:"
    rng.Find.Forward = True
    If rng.Find.Execute Then
        rng.Font.Size = 12
        rng.Font.Bold = True
        rng.Font.Underline = 1
    End If
    
    ' Subtítulos das seções
    Call FormatarTexto(wordDoc, "OPÇÕES DE APOSENTADORIA", 12, True, False)
    Call FormatarTexto(wordDoc, "APOSENTADORIA RÁPIDA", 11, True, False)
    Call FormatarTexto(wordDoc, "APOSENTADORIA SEMI-INTEGRAL (RECOMENDADA)", 11, True, False)
    Call FormatarTexto(wordDoc, "APOSENTADORIA INTEGRAL", 11, True, False)
    Call FormatarTexto(wordDoc, "COMPARATIVO DE GANHOS:", 11, True, False)
    
    ' Formatar valores monetários em negrito
    Call FormatarPadraoValores(wordDoc)
    
    On Error GoTo ErroWord
    
    ' Abrir visualização de impressão automaticamente
    wordApp.PrintPreview = True
    
    ' Mensagem de sucesso
    MsgBox "Documento Word criado com sucesso!" & vbCrLf & vbCrLf & _
           "• Visualização de impressão aberta" & vbCrLf & _
           "• Você pode salvar o documento se desejar" & vbCrLf & _
           "• Feche o Word quando terminar", _
           vbInformation, "Impressão Gerada"
    
    Exit Sub
    
ErroWord:
    MsgBox "Erro ao gerar documento Word:" & vbCrLf & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "Certifique-se de que o Microsoft Word está instalado.", _
           vbCritical, "Erro de Impressão"
End Sub

' ============================================
' Gera cabeçalho com dados da empresa
' ============================================
Private Function GerarCabecalhoImpressao(nomeCliente As String, idCliente As String) As String
    Dim cab As String
    Dim nomeEmpresa As String
    Dim endereco As String
    Dim telefone As String
    Dim email As String
    Dim oab As String
    
    ' Buscar dados da empresa (se existirem)
    nomeEmpresa = GetDadosEmpresa("Nome_Fantasia")
    If nomeEmpresa = "" Then nomeEmpresa = GetDadosEmpresa("Razao_Social")
    
    ' Se não há dados da empresa, não gerar cabeçalho
    If nomeEmpresa = "" Then
        GerarCabecalhoImpressao = ""
        Exit Function
    End If
    
    endereco = GetDadosEmpresa("Endereco")
    telefone = GetDadosEmpresa("Telefone")
    email = GetDadosEmpresa("Email")
    oab = GetDadosEmpresa("OAB_Numero")
    
    ' Montar cabeçalho
    cab = String(57, "=") & vbCrLf
    cab = cab & CentralizarTexto(UCase(nomeEmpresa), 57) & vbCrLf
    
    If endereco <> "" Then
        cab = cab & CentralizarTexto(endereco, 57) & vbCrLf
    End If
    
    If telefone <> "" And email <> "" Then
        cab = cab & CentralizarTexto(telefone & "  |  " & email, 57) & vbCrLf
    ElseIf telefone <> "" Then
        cab = cab & CentralizarTexto(telefone, 57) & vbCrLf
    ElseIf email <> "" Then
        cab = cab & CentralizarTexto(email, 57) & vbCrLf
    End If
    
    If oab <> "" Then
        cab = cab & CentralizarTexto(oab, 57) & vbCrLf
    End If
    
    cab = cab & String(57, "=") & vbCrLf & vbCrLf
    cab = cab & "SIMULAÇÃO DE APOSENTADORIA" & vbCrLf & vbCrLf
    cab = cab & nomeCliente & " - " & idCliente & vbCrLf
    cab = cab & "Data: " & Format(Now, "dd/mm/yyyy HH:nn") & vbCrLf
    cab = cab & String(57, "-") & vbCrLf
    
    GerarCabecalhoImpressao = cab
End Function

' ============================================
' Gera rodapé com assinatura
' ============================================
Private Function GerarRodapeImpressao() As String
    Dim rod As String
    Dim advogado As String
    Dim oab As String
    
    advogado = GetDadosEmpresa("Advogado_Responsavel")
    oab = GetDadosEmpresa("OAB_Numero")
    
    rod = vbCrLf & vbCrLf & String(57, "-") & vbCrLf
    rod = rod & "Relatório gerado pelo Sistema ERP_Prev em " & Format(Now, "dd/mm/yyyy HH:nn") & vbCrLf
    
    If advogado <> "" Then
        rod = rod & "Advogado Responsável: " & advogado
        If oab <> "" Then rod = rod & " - " & oab
        rod = rod & vbCrLf
    End If
    
    rod = rod & vbCrLf & "Este documento é meramente informativo e não substitui análise jurídica completa." & vbCrLf
    rod = rod & String(57, "=")
    
    GerarRodapeImpressao = rod
End Function

' ============================================
' Função auxiliar para formatar texto no Word
' ============================================
Private Sub FormatarTexto(doc As Object, textoFind As String, tamanho As Integer, negrito As Boolean, centralizar As Boolean)
    On Error Resume Next
    Dim rng As Object
    Set rng = doc.Content
    rng.Find.ClearFormatting
    rng.Find.Text = textoFind
    rng.Find.Forward = True
    If rng.Find.Execute Then
        rng.Font.Size = tamanho
        rng.Font.Bold = negrito
        If centralizar Then
            rng.ParagraphFormat.Alignment = 1
        End If
    End If
End Sub

' ============================================
' Função auxiliar para formatar valores monetários
' ============================================
Private Sub FormatarPadraoValores(doc As Object)
    On Error Resume Next
    Dim rng As Object
    Dim i As Integer
    
    For i = 1 To doc.Paragraphs.Count
        Set rng = doc.Paragraphs(i).Range
        If InStr(rng.Text, "Valor estimado:") > 0 Then
            rng.Font.Bold = True
        ElseIf InStr(rng.Text, "Ganho mensal:") > 0 Then
            rng.Font.Bold = True
        End If
    Next i
End Sub

' ============================================
' Função auxiliar para centralizar texto
' ============================================
Private Function CentralizarTexto(texto As String, largura As Integer) As String
    Dim espacos As Integer
    espacos = (largura - Len(texto)) \ 2
    If espacos > 0 Then
        CentralizarTexto = Space(espacos) & texto
    Else
        CentralizarTexto = texto
    End If
End Function

' ============================================
' Botão Voltar
' ============================================
Private Sub cmdVoltar_Click()
    ' Fecha o formulário e libera memória
    Unload Me
End Sub
