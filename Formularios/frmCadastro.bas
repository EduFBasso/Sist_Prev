' ================================================================
' frmCadastro - Formulário de cadastro de clientes
' ================================================================

Private Sub UserForm_Initialize()

    ' Preencher Combo Sexo
    Me.cboSexo.Clear
    Me.cboSexo.AddItem "Masculino"
    Me.cboSexo.AddItem "Feminino"

    ' Preencher Combo Estado Civil
    Me.cboEstadoCivil.Clear
    Me.cboEstadoCivil.AddItem "Solteiro(a)"
    Me.cboEstadoCivil.AddItem "Casado(a)"
    Me.cboEstadoCivil.AddItem "Divorciado(a)"
    Me.cboEstadoCivil.AddItem "Viúvo(a)"
    Me.cboEstadoCivil.AddItem "União Estável"
    
    ' Preencher Combo UF
    Me.cboUF.AddItem "AC"
    Me.cboUF.AddItem "AL"
    Me.cboUF.AddItem "AP"
    Me.cboUF.AddItem "AM"
    Me.cboUF.AddItem "BA"
    Me.cboUF.AddItem "CE"
    Me.cboUF.AddItem "DF"
    Me.cboUF.AddItem "ES"
    Me.cboUF.AddItem "GO"
    Me.cboUF.AddItem "MA"
    Me.cboUF.AddItem "MG"
    Me.cboUF.AddItem "RJ"
    Me.cboUF.AddItem "RS"
    Me.cboUF.AddItem "RO"
    Me.cboUF.AddItem "RR"
    Me.cboUF.AddItem "SC"
    Me.cboUF.AddItem "SP"
    Me.cboUF.AddItem "SE"
    Me.cboUF.AddItem "TO"

    ' Preencher Tipo de Segurado
    Me.cboTipoSegurado.Clear
    Me.cboTipoSegurado.AddItem "Empregado"
    Me.cboTipoSegurado.AddItem "Contribuinte Individual"
    Me.cboTipoSegurado.AddItem "Avulso"
    Me.cboTipoSegurado.AddItem "Doméstico"
    Me.cboTipoSegurado.AddItem "Segurado Especial"
    Me.cboTipoSegurado.AddItem "Facultativo"

    ' Valores padrão
    Me.cboSexo.ListIndex = -1
    Me.cboEstadoCivil.ListIndex = -1
    Me.cboTipoSegurado.ListIndex = -1

    ' ID só é preenchido quando chamamos CarregarCliente
    Me.txtID.Value = ""

    ' Idade apenas leitura
    On Error Resume Next
    Me.txtIdade.Value = ""
    Me.txtIdade.Locked = True
    On Error GoTo 0

End Sub

'======================================================
' Limpar Formulário - Apaga os campos deste formulário
'======================================================

Private Sub LimparFormulario()

    ' Campos técnicos
    Me.txtID.Value = ""

    ' Dados pessoais
    Me.txtNome.Value = ""
    Me.txtCPF.Value = ""
    Me.txtPIS.Value = ""
    Me.txtRG.Value = ""
    Me.txtOrgao.Value = ""
    Me.txtNascimento.Value = ""
    Me.cboSexo.ListIndex = -1
    Me.cboEstadoCivil.ListIndex = -1
    Me.txtFiliacao.Value = ""

    ' Contato
    Me.txtTelefone.Value = ""
    Me.txtCelular.Value = ""
    Me.txtEmail.Value = ""

    ' Endereço
    Me.txtCEP.Value = ""
    Me.txtEndereco.Value = ""
    Me.txtNumero.Value = ""
    Me.txtComplemento.Value = ""
    Me.txtBairro.Value = ""
    Me.txtCidade.Value = ""
    Me.cboUF.ListIndex = -1

    ' Tipo de segurado
    Me.cboTipoSegurado.ListIndex = -1

    ' Condições especiais
    Me.chkPossuiEspecial.Value = False
    Me.chkPossuiRural.Value = False
    Me.chkPossuiMilitar.Value = False
    Me.chkPossuiExterior.Value = False
    Me.chkPossuiConcomitante.Value = False
    Me.chkPossuiAtraso.Value = False
    Me.chkPossuiComplementar.Value = False

    ' Observações
    Me.txtObservacoes.Value = ""

    ' Vínculos
    Me.lstVinculos.Clear

    On Error Resume Next
    Me.txtIdade.Value = ""
    On Error GoTo 0

End Sub

Public Sub NovoCliente()
    Call LimparFormulario
End Sub

Private Sub txtNascimento_Change()

    On Error GoTo Fim

    If IsDate(Me.txtNascimento.Value) Then
        Me.txtIdade.Value = CalcularIdade(CDate(Me.txtNascimento.Value)) & " anos"
    Else
        Me.txtIdade.Value = ""
    End If

Fim:
End Sub

Private Sub cmdNovoVinculo_Click()

    If Me.txtID.Value = "" Then
        MsgBox "Carregue um cliente antes de adicionar vínculos.", vbExclamation
        Exit Sub
    End If

    ' Limpa o formulário de vínculos
    frmVinculos.txtIDVinculo.Value = ""
    frmVinculos.txtInicio.Value = ""
    frmVinculos.txtFim.Value = ""
    frmVinculos.cboTipo.ListIndex = -1
    frmVinculos.cboEspecial.Value = "Não"
    frmVinculos.cboGrau.Value = ""
    frmVinculos.cboGrau.Enabled = False
    frmVinculos.txtSalario.Value = ""
    frmVinculos.txtObs.Value = ""

    ' Define o ID do cliente
    frmVinculos.txtIDCliente.Value = Me.txtID.Value

    ' Abre o formulário de vínculos
    frmVinculos.Show

End Sub

' Fecha o formulário principal
Private Sub cmdVoltar_Click()
    Unload Me
    ' frmPrincipal.Show
End Sub

' Abrir formulário de busca
Private Sub cmdBuscar_Click()
    frmBusca.Show
End Sub

' Btn Excluir
Private Sub cmdExcluir_Click()

    If Me.txtID.Value = "" Then
        MsgBox "Nenhum cliente selecionado para excluir.", vbExclamation
        Exit Sub
    End If

    Call ExcluirCliente(Me.txtID.Value)

    MsgBox "Cliente excluído com sucesso!", vbInformation

    Call LimparFormulario

End Sub

' Btn Novo
Private Sub cmdNovo_Click()
    Call LimparFormulario
    Me.txtNome.SetFocus
End Sub

' Btn Salvar
Private Sub cmdSalvar_Click()

    If Me.txtID.Value = "" Then
        Me.txtID.Value = GerarNovoID()
    End If
    
    If Trim(Me.txtNome.Value) = "" Then
        MsgBox "Informe o nome do cliente.", vbExclamation
        Exit Sub
    End If

    Dim dados As Collection
    Set dados = ColetarDados

    ' Envia para o módulo de cadastro
    Call SalvarCliente(dados)

    MsgBox "Registro salvo com sucesso!", vbInformation

End Sub

' ColetarDados - Transforma todos os campos em uma coleção organizada.
Function ColetarDados() As Collection

    Dim c As New Collection

    ' ID
    c.Add IIf(Me.txtID.Value = "", 0, CLng(Me.txtID.Value)), "ID"

    ' Dados pessoais
    c.Add Me.txtNome.Value, "Nome"
    c.Add Me.txtCPF.Value, "CPF"
    c.Add Me.txtPIS.Value, "PIS"
    c.Add Me.txtRG.Value, "RG"
    c.Add Me.txtOrgao.Value, "Orgao"
    c.Add Me.txtNascimento.Value, "Nascimento"
    c.Add Me.cboSexo.Value, "Sexo"
    c.Add Me.cboEstadoCivil.Value, "EstadoCivil"
    c.Add Me.txtFiliacao.Value, "Filiacao"

    ' Contato
    c.Add Me.txtTelefone.Value, "Telefone"
    c.Add Me.txtCelular.Value, "Celular"
    c.Add Me.txtEmail.Value, "Email"

    ' Endereço
    c.Add Me.txtCEP.Value, "CEP"
    c.Add Me.txtEndereco.Value, "Endereco"
    c.Add Me.txtNumero.Value, "Numero"
    c.Add Me.txtComplemento.Value, "Complemento"
    c.Add Me.txtBairro.Value, "Bairro"
    c.Add Me.txtCidade.Value, "Cidade"
    c.Add Me.cboUF.Value, "UF"

    ' Tipo de segurado
    c.Add Me.cboTipoSegurado.Value, "TipoSegurado"

    ' Condições especiais
    c.Add Me.chkPossuiEspecial.Value, "PossuiEspecial"
    c.Add Me.chkPossuiRural.Value, "PossuiRural"
    c.Add Me.chkPossuiMilitar.Value, "PossuiMilitar"
    c.Add Me.chkPossuiExterior.Value, "PossuiExterior"
    c.Add Me.chkPossuiConcomitante.Value, "PossuiConcomitante"
    c.Add Me.chkPossuiAtraso.Value, "PossuiAtraso"
    c.Add Me.chkPossuiComplementar.Value, "PossuiComplementar"

    ' Observações
    c.Add Me.txtObservacoes.Value, "Observacoes"

    Set ColetarDados = c

End Function

' ================================================================
' ExcluirCliente - Remove o registro da planilha Cadastro_Clientes
' ================================================================

Sub ExcluirCliente(ID As Long)

    Dim ws As Worksheet
    Dim linha As Long

    Set ws = Sheets("Cadastro_Clientes")
    linha = BuscarLinhaCliente(ID)

    If linha = 0 Then
        MsgBox "Cliente não encontrado para exclusão.", vbExclamation
        Exit Sub
    End If

    ' Limpa a linha inteira
    ws.Rows(linha).ClearContents

End Sub

Private Sub lstVinculos_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    If Me.lstVinculos.ListIndex = -1 Then Exit Sub

    Dim ID_Vinculo As Long
    ID_Vinculo = Me.lstVinculos.List(Me.lstVinculos.ListIndex, 0)

    Call CarregarVinculo(ID_Vinculo)

    frmVinculos.Show

End Sub

Private Sub cmdSimular_Click()

    ' Falta criar o código para iniciar a simulação no frmSimulacoes
End Sub
