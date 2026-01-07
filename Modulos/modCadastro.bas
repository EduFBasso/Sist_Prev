' ============================================
' modCadastro - Operações do cadastro
' ============================================

Option Explicit

Sub SalvarCliente(dados As Collection)

    Dim ws As Worksheet
    Dim linha As Long

    Set ws = Sheets("Cadastro_Clientes")

    ' Tenta localizar o ID na planilha
    linha = BuscarLinhaCliente(dados("ID"))

    ' Se não encontrou, é novo cliente
    If linha = 0 Then
        linha = ProximaLinha(ws)
    End If

    ' Agora salva os dados
    ws.Cells(linha, 1).Value = dados("ID")
    ws.Cells(linha, 2).Value = dados("Nome")
    ws.Cells(linha, 3).Value = dados("CPF")
    ws.Cells(linha, 4).Value = dados("PIS")
    ws.Cells(linha, 5).Value = dados("RG")
    ws.Cells(linha, 6).Value = dados("Orgao")
    ws.Cells(linha, 7).Value = dados("Nascimento")
    ws.Cells(linha, 8).Value = dados("Sexo")
    ws.Cells(linha, 9).Value = dados("EstadoCivil")
    ws.Cells(linha, 10).Value = dados("Telefone")
    ws.Cells(linha, 11).Value = dados("Celular")
    ws.Cells(linha, 12).Value = dados("Email")
    ws.Cells(linha, 13).Value = dados("CEP")
    ws.Cells(linha, 14).Value = dados("Endereco")
    ws.Cells(linha, 15).Value = dados("Numero")
    ws.Cells(linha, 16).Value = dados("Complemento")
    ws.Cells(linha, 17).Value = dados("Bairro")
    ws.Cells(linha, 18).Value = dados("Cidade")
    ws.Cells(linha, 19).Value = dados("UF")
    ws.Cells(linha, 20).Value = dados("Filiacao")
    ws.Cells(linha, 21).Value = dados("TipoSegurado")
    ws.Cells(linha, 22).Value = dados("Especial")
    ws.Cells(linha, 23).Value = dados("Rural")
    ws.Cells(linha, 24).Value = dados("Militar")
    ws.Cells(linha, 25).Value = dados("Exterior")
    ws.Cells(linha, 26).Value = dados("Concomitante")
    ws.Cells(linha, 27).Value = dados("Atraso")
    ws.Cells(linha, 28).Value = dados("Complementar")
    ws.Cells(linha, 29).Value = dados("Observacoes")

    ' Idade calculada (opcional - última coluna livre)
    On Error Resume Next
    If IsDate(dados("Nascimento")) Then
        ws.Cells(linha, 30).Value = CalcularIdade(CDate(dados("Nascimento")))
    Else
        ws.Cells(linha, 30).Value = ""
    End If
    On Error GoTo 0

End Sub

' ===========================================================================================================
' BuscarLinhaCliente - procura o ID do cliente na planilha Cadastro_Clientes e retorna a linha onde ele está
' ===========================================================================================================

Function BuscarLinhaCliente(ID As Long) As Long
    Dim ws As Worksheet
    Dim ultima As Long
    Dim i As Long

    Set ws = Sheets("Cadastro_Clientes")
    ultima = ws.Cells(ws.Rows.count, 1).End(xlUp).Row

    For i = 2 To ultima
        If ws.Cells(i, 1).Value = ID Then
            BuscarLinhaCliente = i
            Exit Function
        End If
    Next i

    BuscarLinhaCliente = 0 ' Não encontrado
End Function

' =======================================================================================
' CarregarCliente - recebe um ID, busca a linha e preenche todos os campos do frmCadastro
' =======================================================================================

Sub CarregarCliente(ID As Long)

    Dim ws As Worksheet
    Dim linha As Long

    Set ws = Sheets("Cadastro_Clientes")
    linha = BuscarLinhaCliente(ID)

    If linha = 0 Then
        MsgBox "Cliente não encontrado.", vbExclamation
        Exit Sub
    End If

    ' Campos técnicos
    frmCadastro.txtID.Value = ws.Cells(linha, 1).Value

    ' Dados pessoais
    frmCadastro.txtNome.Value = ws.Cells(linha, 2).Value
    frmCadastro.txtCPF.Value = ws.Cells(linha, 3).Value
    frmCadastro.txtPIS.Value = ws.Cells(linha, 4).Value
    frmCadastro.txtRG.Value = ws.Cells(linha, 5).Value
    frmCadastro.txtOrgao.Value = ws.Cells(linha, 6).Value
    frmCadastro.txtNascimento.Value = ws.Cells(linha, 7).Value
    On Error Resume Next
    If IsDate(ws.Cells(linha, 7).Value) Then
        frmCadastro.txtIdade.Value = CalcularIdade(CDate(ws.Cells(linha, 7).Value)) & " anos"
    Else
        frmCadastro.txtIdade.Value = ""
    End If
    On Error GoTo 0
    frmCadastro.cboSexo.Value = ws.Cells(linha, 8).Value
    frmCadastro.cboEstadoCivil.Value = ws.Cells(linha, 9).Value

    ' Contato
    frmCadastro.txtTelefone.Value = ws.Cells(linha, 10).Value
    frmCadastro.txtCelular.Value = ws.Cells(linha, 11).Value
    frmCadastro.txtEmail.Value = ws.Cells(linha, 12).Value

    ' Endereço
    frmCadastro.txtCEP.Value = ws.Cells(linha, 13).Value
    frmCadastro.txtEndereco.Value = ws.Cells(linha, 14).Value
    frmCadastro.txtNumero.Value = ws.Cells(linha, 15).Value
    frmCadastro.txtComplemento.Value = ws.Cells(linha, 16).Value
    frmCadastro.txtBairro.Value = ws.Cells(linha, 17).Value
    frmCadastro.txtCidade.Value = ws.Cells(linha, 18).Value
    frmCadastro.cboUF.Value = ws.Cells(linha, 19).Value
    frmCadastro.txtFiliacao.Value = ws.Cells(linha, 20).Value
    frmCadastro.cboTipoSegurado.Value = ws.Cells(linha, 21).Value
    
    ' Características dos Vínculos
    frmCadastro.chkPossuiEspecial.Value = ws.Cells(linha, 22).Value
    frmCadastro.chkPossuiRural.Value = ws.Cells(linha, 23).Value
    frmCadastro.chkPossuiMilitar.Value = ws.Cells(linha, 24).Value
    frmCadastro.chkPossuiExterior.Value = ws.Cells(linha, 25).Value
    frmCadastro.chkPossuiConcomitante.Value = ws.Cells(linha, 26).Value
    frmCadastro.chkPossuiAtraso.Value = ws.Cells(linha, 27).Value
    frmCadastro.chkPossuiComplementar.Value = ws.Cells(linha, 28).Value
    
    ' Carregar vínculos do cliente
    Call CarregarVinculosCliente(ID)

End Sub

Function GerarNovoID() As Long

    Dim ws As Worksheet
    Dim ultima As Long

    Set ws = Sheets("Cadastro_Clientes")

    ultima = ws.Cells(ws.Rows.count, 1).End(xlUp).Row

    If ultima < 2 Then
        GerarNovoID = 1
    Else
        GerarNovoID = ws.Cells(ultima, 1).Value + 1
    End If

End Function
