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
    ' O arquivo selecionado é *_vinculos_estruturado.csv
    ' O arquivo de dados é *_dados_cliente.csv (mesmo prefixo base)
    Dim caminhoBase As String
    caminhoBase = CStr(caminho)
    ' Remove "_vinculos_estruturado.csv" se presente
    If InStr(caminhoBase, "_vinculos_estruturado.csv") > 0 Then
        caminhoBase = Replace(caminhoBase, "_vinculos_estruturado.csv", "")
    Else
        ' Remove extensão .csv
        caminhoBase = Left$(caminhoBase, InStrRev(caminhoBase, ".") - 1)
    End If
    caminhoDados = caminhoBase & "_dados_cliente.csv"

    ' 1. Tenta carregar dados do cliente a partir do CSV de cabeçalho
    Dim dadosCliente As Collection
    Set dadosCliente = Nothing

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
                    Set dadosCliente = New Collection
                    dadosCliente.Add novoID, "ID"
                    dadosCliente.Add partes(2), "Nome"
                    dadosCliente.Add partes(1), "CPF"
                    dadosCliente.Add partes(0), "PIS"
                    dadosCliente.Add "", "RG"
                    dadosCliente.Add "", "Orgao"
                    dadosCliente.Add partes(3), "Nascimento"
                    dadosCliente.Add "", "Sexo"
                    dadosCliente.Add "", "EstadoCivil"
                    dadosCliente.Add "", "Telefone"
                    dadosCliente.Add "", "Celular"
                    dadosCliente.Add "", "Email"
                    dadosCliente.Add "", "CEP"
                    dadosCliente.Add "", "Endereco"
                    dadosCliente.Add "", "Numero"
                    dadosCliente.Add "", "Complemento"
                    dadosCliente.Add "", "Bairro"
                    dadosCliente.Add "", "Cidade"
                    dadosCliente.Add "", "UF"
                    dadosCliente.Add partes(4), "Filiacao"
                    dadosCliente.Add "", "TipoSegurado"
                    dadosCliente.Add "Não", "Especial"
                    dadosCliente.Add False, "Rural"
                    dadosCliente.Add False, "Militar"
                    dadosCliente.Add False, "Exterior"
                    dadosCliente.Add False, "Concomitante"
                    dadosCliente.Add False, "Atraso"
                    dadosCliente.Add False, "Complementar"
                    dadosCliente.Add "Importado do CNIS", "Observacoes"
                End If
            End If
        End If
FimLeitura:
        On Error Resume Next
        Close #fNum
        On Error GoTo 0
    End If

    ' 2. Salva o cliente na planilha antes de abrir o formulário
    If Not dadosCliente Is Nothing Then
        Call SalvarCliente(dadosCliente)
    Else
        ' Se não houver dados, cria cliente vazio
        Set dadosCliente = New Collection
        dadosCliente.Add novoID, "ID"
        dadosCliente.Add "", "Nome"
        dadosCliente.Add "", "CPF"
        dadosCliente.Add "", "PIS"
        dadosCliente.Add "", "RG"
        dadosCliente.Add "", "Orgao"
        dadosCliente.Add "", "Nascimento"
        dadosCliente.Add "", "Sexo"
        dadosCliente.Add "", "EstadoCivil"
        dadosCliente.Add "", "Telefone"
        dadosCliente.Add "", "Celular"
        dadosCliente.Add "", "Email"
        dadosCliente.Add "", "CEP"
        dadosCliente.Add "", "Endereco"
        dadosCliente.Add "", "Numero"
        dadosCliente.Add "", "Complemento"
        dadosCliente.Add "", "Bairro"
        dadosCliente.Add "", "Cidade"
        dadosCliente.Add "", "UF"
        dadosCliente.Add "", "Filiacao"
        dadosCliente.Add "", "TipoSegurado"
        dadosCliente.Add "Não", "Especial"
        dadosCliente.Add False, "Rural"
        dadosCliente.Add False, "Militar"
        dadosCliente.Add False, "Exterior"
        dadosCliente.Add False, "Concomitante"
        dadosCliente.Add False, "Atraso"
        dadosCliente.Add False, "Complementar"
        dadosCliente.Add "Importado do CNIS", "Observacoes"
        Call SalvarCliente(dadosCliente)
    End If

    ' 3. Carrega o formulário (inicializa sem mostrar), preenche os dados e então exibe
    Load frmCadastro
    Call CarregarCliente(novoID)
    
    MsgBox "Vínculos importados com sucesso para o cliente ID " & novoID & ". Complete os dados cadastrais.", _
           vbInformation
    
    frmCadastro.Show

End Sub
