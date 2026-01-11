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
    
    If nascimento = 0 Then
    On Error Resume Next
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
