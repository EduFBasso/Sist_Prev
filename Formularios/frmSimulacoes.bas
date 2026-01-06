Private Sub PreencherDadosIniciais(ID_Cliente As Long)

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
    nascimento = ws.Cells(linha, 7).Value

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

    ID = CLng(Me.Tag)

    ' Verifica qual regra foi escolhida
    If Me.optTempoContribuicao.Value = True Then
        Set r = RegraTempoContribuicao(ID)
    Else
        MsgBox "Selecione uma regra.", vbExclamation
        Exit Sub
    End If

    ' Preencher resultados
    Me.txtDireito.Value = r("Direito")
    Me.txtFalta.Value = Format(r("Falta"), "0.00") & " anos"
    Me.txtIdadeProj.Value = r("IdadeProjetada")
    Me.txtDataProvavel.Value = r("DataPrevista")
    Me.txtObs.Value = r("Obs")

End Sub