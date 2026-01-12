'==================================================================
' Formulário de Vínculos
'==================================================================

Private Sub UserForm_Initialize()

    ' Tipo de vínculo
    With Me.cboTipo
        .Clear
        .AddItem "Empregado"
        .AddItem "Contribuinte Individual"
        .AddItem "Avulso"
        .AddItem "Doméstico"
        .AddItem "Segurado Especial"
        .AddItem "Servidor Público"
        .AddItem "Militar"
    End With

    ' Especial
    With Me.cboEspecial
        .Clear
        .AddItem "Sim"
        .AddItem "Não"
    End With
    
    ' Valor padrão ao executar o form
    Me.cboEspecial.Value = "Não"
    Me.cboGrau.Enabled = False
    
    ' Desativando check box
    Me.chkRural.Value = False
    Me.chkMilitar.Value = False
    Me.chkExterior.Value = False
    Me.chkConcomitante.Value = False
    Me.chkAtraso.Value = False
    Me.chkComplementar.Value = False

    ' Grau (somente se Especial = Sim)
    With Me.cboGrau
        .Clear
        .AddItem "15"
        .AddItem "20"
        .AddItem "25"
    End With

End Sub

Private Sub LimparFormulario()

    Me.txtIDVinculo.Value = ""
    Me.txtInicio.Value = ""
    Me.txtFim.Value = ""
    Me.cboTipo.ListIndex = -1
    Me.cboEspecial.ListIndex = -1
    Me.cboGrau.ListIndex = -1
    Me.txtSalario.Value = ""
    Me.txtObs.Value = ""

End Sub


Private Function ColetarDados() As Collection

    Dim c As New Collection
    Dim idV As Long

    ' Trata ID_Vinculo corretamente
    If Trim(Me.txtIDVinculo.Value) = "" Then
        idV = 0
    Else
        idV = CLng(Me.txtIDVinculo.Value)
    End If

    c.Add idV, "ID_Vinculo"
    c.Add CLng(Me.txtIDCliente.Value), "ID_Cliente"
    c.Add Me.txtInicio.Value, "Inicio"
    c.Add Me.txtFim.Value, "Fim"
    c.Add Me.cboTipo.Value, "Tipo"
    c.Add Me.cboEspecial.Value, "Especial"
    c.Add Me.cboGrau.Value, "Grau"
    c.Add Me.txtSalario.Value, "Salario"
    c.Add Me.txtObs.Value, "Observacoes"
    c.Add Me.chkRural.Value, "Rural"
    c.Add Me.chkMilitar.Value, "Militar"
    c.Add Me.chkExterior.Value, "Exterior"
    c.Add Me.chkConcomitante.Value, "Concomitante"
    c.Add Me.chkAtraso.Value, "Atraso"
    c.Add Me.chkComplementar.Value, "Complementar"

    Set ColetarDados = c

End Function


Private Sub cmdSalvar_Click()

    Dim dados As Collection
    Set dados = ColetarDados

    Call SalvarVinculo(dados)

    MsgBox "Vínculo salvo com sucesso!", vbInformation
    
    Call CarregarVinculosCliente(Me.txtIDCliente.Value)
    Call LimparFormulario
    Me.txtInicio.SetFocus

End Sub


Private Sub cmdNovo_Click()
    Call LimparFormulario
    Me.txtInicio.SetFocus
End Sub


Private Sub cmdExcluir_Click()

    If Me.txtIDVinculo.Value = "" Then
        MsgBox "Nenhum vínculo selecionado para excluir.", vbExclamation
        Exit Sub
    End If

    Call ExcluirVinculo(Me.txtIDVinculo.Value)

    MsgBox "Vínculo excluído com sucesso!", vbInformation
    Call LimparFormulario
    
    Call CarregarVinculosCliente(Me.txtIDCliente.Value)

End Sub


Private Sub cmdVoltar_Click()
    Unload Me
End Sub


Private Sub cboEspecial_Change()
    If Me.cboEspecial.Value = "Sim" Then
        Me.cboGrau.Enabled = True
    Else
        Me.cboGrau.Enabled = False
        Me.cboGrau.Value = ""
    End If
End Sub

