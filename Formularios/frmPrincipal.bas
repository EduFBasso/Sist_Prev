' ================================================================
' frmPrincipal - Formulário principal do sistema
' ================================================================

Option Explicit

Private Sub cmdBuscarCliente_Click()
    frmBusca.Show
End Sub

Private Sub cmdCadastro_Click()
    frmCadastro.NovoCliente
    frmCadastro.Show
End Sub

Private Sub cmdDocumentos_Click()
    frmDocumentos.Show
End Sub

Private Sub cmdSimulacoes_Click()
    frmSimulacoes.Show
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdImportarCNIS_Click()

    Dim caminho As Variant
    Dim novoID As Long
    Dim caminhoDados As String
    Dim fNum As Integer
    Dim linha As String
    Dim partes() As String

    ' Seleciona o arquivo CSV de vínculos estruturados
    caminho = Application.GetOpenFilename("Arquivos CSV (*.csv),*.csv", , _
                                          "Selecione o arquivo de vínculos do CNIS")
    If caminho = False Then Exit Sub

    ' Gera um novo ID de cliente baseado na planilha Cadastro_Clientes
    novoID = GerarNovoID()

    ' Importa vínculos para esse novo cliente
    Call ImportarVinculosDeCSV(CStr(caminho), novoID)

    ' Tenta carregar dados do cliente a partir do CSV de cabeçalho
    caminhoDados = Left$(caminho, InStrRev(caminho, ".") - 1) & "_dados_cliente.csv"

    frmCadastro.NovoCliente
    frmCadastro.txtID.Value = novoID

    If Dir(caminhoDados) <> "" Then

        fNum = FreeFile
        On Error GoTo FimLeitura
        Open caminhoDados For Input As #fNum

        ' Pula cabeçalho
        Line Input #fNum, linha

        If Not EOF(fNum) Then
            Line Input #fNum, linha
            If Trim$(linha) <> "" Then
                partes = Split(linha, ";")

                ' Esperado: NIT;CPF;Nome;DataNascimento;NomeMae
                If UBound(partes) >= 4 Then
                    If frmCadastro.txtPIS.Enabled Then frmCadastro.txtPIS.Value = partes(0)
                    frmCadastro.txtCPF.Value = partes(1)
                    frmCadastro.txtNome.Value = partes(2)
                    frmCadastro.txtNascimento.Value = partes(3)
                    frmCadastro.txtFiliacao.Value = partes(4)
                End If
            End If
        End If

FimLeitura:
        On Error Resume Next
        Close #fNum
        On Error GoTo 0

    End If

    ' Abre o cadastro para o usuário revisar/completar dados
    frmCadastro.NovoCliente
    frmCadastro.txtID.Value = novoID
    frmCadastro.Show

    MsgBox "Vínculos importados com sucesso para o cliente ID " & novoID & ". Complete os dados cadastrais.", _
           vbInformation

End Sub
