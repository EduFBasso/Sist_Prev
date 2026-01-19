' ================================================================
' frmPrincipal - Formulário principal do sistema V1
' ================================================================

Option Explicit

Private Sub cmdBuscarCliente_Click()
    ' Validar se há clientes cadastrados ANTES de abrir o formulário
    Dim ws As Worksheet
    Dim ultima As Long
    
    On Error Resume Next
    Set ws = Sheets("Cadastro_Clientes")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Planilha 'Cadastro_Clientes' não encontrada." & vbCrLf & _
               "O sistema precisa dessa planilha para funcionar.", vbCritical, "Erro"
        Exit Sub
    End If
    
    ultima = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    ' Se não há clientes cadastrados (apenas cabeçalho na linha 1)
    If ultima < 2 Then
        MsgBox "⚠️  Nenhum cliente cadastrado!" & vbCrLf & vbCrLf & _
               "Para começar:" & vbCrLf & _
               "1. Use o botão 'Novo Cliente' para cadastrar manualmente, OU" & vbCrLf & _
               "2. Importe dados do CNIS usando 'Importar CNIS'", _
               vbInformation, "Cadastro Vazio"
        Exit Sub
    End If
    
    ' Se há clientes, abrir o formulário de busca
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
    ' Usa a nova função automática que importa tudo de uma vez
    ' (dados do cliente + vínculos + remunerações)
    Call ImportarCNIS_Automatico
End Sub

Private Sub cmdAtualizarINPC_Click()
    ' Atualizar índices INPC do Banco Central
    Dim sucesso As Boolean
    
    sucesso = AtualizarIndicesINPC()
    
    If sucesso Then
        ' Atualizar label com a data
        Call AtualizarLabelDataInpcSelic
        
        MsgBox "Índices INPC atualizados com sucesso!" & vbCrLf & vbCrLf & _
               "Data: " & GetParametro("Data_Atualizacao_INPC") & vbCrLf & vbCrLf & _
               "Os cálculos de simulação já estão usando os novos fatores de correção monetária.", _
               vbInformation, "Atualização Concluída"
    Else
        MsgBox "Erro ao atualizar índices INPC." & vbCrLf & _
               "Verifique sua conexão com a internet e tente novamente.", _
               vbCritical, "Erro"
    End If
End Sub

Private Sub UserForm_Initialize()
    ' Atualizar label da data ao abrir o formulário
    Call AtualizarLabelDataInpcSelic

    ' Inicializar seleção do índice de correção (INPC/SELIC)
    Call InicializarIndiceCorrecao
End Sub

' ================================================================
' Índice de correção monetária (INPC x SELIC)
' - Persistência: Config_Regras / parâmetro "Indice_Correcao_Remuneracoes"
' ================================================================
Private Sub InicializarIndiceCorrecao()
    Dim indiceAtual As String

    indiceAtual = UCase(Trim(CStr(GetParametro("Indice_Correcao_Remuneracoes"))))

    ' Se não existir, cria com padrão INPC
    If indiceAtual = "" Or indiceAtual = "0" Then
        Call SetParametro("Indice_Correcao_Remuneracoes", "INPC", _
                          "Indice de correcao das remuneracoes para media (INPC/SELIC)")
        indiceAtual = "INPC"
    End If

    ' Preencher ComboBox (se existir no layout)
    If ControleExiste("cboIndiceCorrecao") Then
        Dim cbo As Object
        Set cbo = Me.Controls("cboIndiceCorrecao")
        cbo.Clear
        cbo.AddItem "INPC"
        cbo.AddItem "SELIC"
        cbo.Value = indiceAtual
    End If

    ' Preencher OptionButtons (se existirem no layout)
    If ControleExiste("optIndiceINPC") Then
        Me.Controls("optIndiceINPC").Value = (indiceAtual = "INPC")
    End If
    If ControleExiste("optIndiceSELIC") Then
        Me.Controls("optIndiceSELIC").Value = (indiceAtual = "SELIC")
    End If
End Sub

Private Function ControleExiste(ByVal nomeControle As String) As Boolean
    On Error GoTo NaoExiste
    Dim tmp As Object
    Set tmp = Me.Controls(nomeControle)
    ControleExiste = True
    Exit Function
NaoExiste:
    ControleExiste = False
End Function

Private Sub SalvarIndiceCorrecao(ByVal indice As String)
    indice = UCase(Trim(indice))
    If indice <> "INPC" And indice <> "SELIC" Then Exit Sub
    Call SetParametro("Indice_Correcao_Remuneracoes", indice, _
                      "Indice de correcao das remuneracoes para media (INPC/SELIC)")
End Sub

' --- Eventos (escolha 1): ComboBox ---
Private Sub cboIndiceCorrecao_Change()
    ' Se o ComboBox não existir no layout, não faz nada (evita erro de compilação no Windows)
    If Not ControleExiste("cboIndiceCorrecao") Then Exit Sub
    Call SalvarIndiceCorrecao(CStr(Me.Controls("cboIndiceCorrecao").Value))
End Sub

' --- Eventos (escolha 2): OptionButtons ---
Private Sub optIndiceINPC_Click()
    If Me.optIndiceINPC.Value = True Then Call SalvarIndiceCorrecao("INPC")
End Sub

Private Sub optIndiceSELIC_Click()
    If Me.optIndiceSELIC.Value = True Then Call SalvarIndiceCorrecao("SELIC")
End Sub

Private Sub AtualizarLabelDataInpcSelic()
    ' Atualiza o label com a última data de atualização
    Dim dataAtual As String
    
    dataAtual = GetParametro("Data_Atualizacao_INPC")
    
    If dataAtual <> "" And Not IsNull(dataAtual) Then
        Me.lblDataAtualizacao.Caption = "Última atualização INPC: " & dataAtual
        Me.lblDataAtualizacao.ForeColor = RGB(0, 128, 0)  ' Verde
    Else
        Me.lblDataAtualizacao.Caption = "INPC não atualizado - clique no botão acima"
        Me.lblDataAtualizacao.ForeColor = RGB(255, 0, 0)  ' Vermelho
    End If
End Sub
