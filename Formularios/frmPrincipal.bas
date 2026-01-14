' ================================================================
' frmPrincipal - Formulário principal do sistema
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
        Call AtualizarLabelDataINPC
        
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
    Call AtualizarLabelDataINPC
End Sub

Private Sub AtualizarLabelDataINPC()
    ' Atualiza o label com a última data de atualização
    Dim dataAtual As String
    
    dataAtual = GetParametro("Data_Atualizacao_INPC")
    
    If dataAtual <> "" And Not IsNull(dataAtual) Then
        Me.lblDataINPC.Caption = "Última atualização INPC: " & dataAtual
        Me.lblDataINPC.ForeColor = RGB(0, 128, 0)  ' Verde
    Else
        Me.lblDataINPC.Caption = "INPC não atualizado - clique no botão acima"
        Me.lblDataINPC.ForeColor = RGB(255, 0, 0)  ' Vermelho
    End If
End Sub
