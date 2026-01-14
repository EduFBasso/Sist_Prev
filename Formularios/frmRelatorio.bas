'------------------------------------------------------------------
' Formulário: frmRelatorio
' Descrição: Formulário para visualização e impressão de relatório
'            de simulação de aposentadoria (tamanho A4)
'------------------------------------------------------------------

Option Explicit

Private mID_Cliente As Long
Private mNomeCliente As String
Private mResultados As String

' ============================================
' Método público para inicializar o relatório
' ============================================
Public Sub GerarRelatorio(ID_Cliente As Long, nomeCliente As String, txtResultados As String)
    mID_Cliente = ID_Cliente
    mNomeCliente = nomeCliente
    mResultados = txtResultados
    
    Call PreencherRelatorio
End Sub

' ============================================
' Preenche os campos do relatório
' ============================================
Private Sub PreencherRelatorio()
    Dim cabecalho As String
    Dim rodape As String
    Dim dataHora As String
    
    ' Dados da empresa
    Dim nomeEmpresa As String
    Dim endereco As String
    Dim telefone As String
    Dim email As String
    Dim oab As String
    Dim advogado As String
    
    nomeEmpresa = GetDadosEmpresa("Nome_Fantasia")
    If nomeEmpresa = "" Then nomeEmpresa = GetDadosEmpresa("Razao_Social")
    If nomeEmpresa = "" Then nomeEmpresa = "ADVOCACIA PREVIDENCIÁRIA"
    
    endereco = GetDadosEmpresa("Endereco")
    telefone = GetDadosEmpresa("Telefone")
    email = GetDadosEmpresa("Email")
    oab = GetDadosEmpresa("OAB_Numero")
    advogado = GetDadosEmpresa("Advogado_Responsavel")
    
    ' Data e hora atual
    dataHora = Format(Now, "dd/mm/yyyy HH:nn")
    
    ' Montar cabeçalho
    cabecalho = String(80, "=") & vbCrLf
    cabecalho = cabecalho & Space(25) & UCase(nomeEmpresa) & vbCrLf
    
    If endereco <> "" Then
        cabecalho = cabecalho & Space(20) & endereco & vbCrLf
    End If
    
    If telefone <> "" Or email <> "" Then
        cabecalho = cabecalho & Space(15) & telefone & "  |  " & email & vbCrLf
    End If
    
    If oab <> "" Then
        cabecalho = cabecalho & Space(30) & oab & vbCrLf
    End If
    
    cabecalho = cabecalho & String(80, "=") & vbCrLf & vbCrLf
    cabecalho = cabecalho & "SIMULAÇÃO DE APOSENTADORIA" & vbCrLf
    cabecalho = cabecalho & "Cliente: " & mNomeCliente & " (ID: " & mID_Cliente & ")" & vbCrLf
    cabecalho = cabecalho & "Data: " & dataHora & vbCrLf
    cabecalho = cabecalho & String(80, "-") & vbCrLf & vbCrLf
    
    ' Montar rodapé
    rodape = vbCrLf & vbCrLf & String(80, "-") & vbCrLf
    rodape = rodape & "Relatório gerado pelo Sistema ERP_Prev em " & dataHora & vbCrLf
    
    If advogado <> "" Then
        rodape = rodape & "Advogado Responsável: " & advogado
        If oab <> "" Then rodape = rodape & " - " & oab
        rodape = rodape & vbCrLf
    End If
    
    rodape = rodape & vbCrLf & "Este documento é meramente informativo e não substitui análise jurídica completa." & vbCrLf
    rodape = rodape & String(80, "=")
    
    ' Preencher campos do formulário
    Me.txtCabecalho.Value = cabecalho
    Me.txtConteudo.Value = mResultados
    Me.txtRodape.Value = rodape
End Sub

' ============================================
' Botão Imprimir
' ============================================
Private Sub cmdImprimir_Click()
    ' Abrir diálogo de impressão
    Me.PrintForm
End Sub

' ============================================
' Botão Exportar PDF
' ============================================
Private Sub cmdExportarPDF_Click()
    Dim caminhoExcel As String
    Dim nomeArquivo As String
    Dim caminhoCompleto As String
    
    caminhoExcel = ThisWorkbook.Path
    nomeArquivo = "Simulacao_" & mID_Cliente & "_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
    caminhoCompleto = caminhoExcel & Application.PathSeparator & "saida" & _
                      Application.PathSeparator & nomeArquivo
    
    ' Criar pasta saida se não existir
    On Error Resume Next
    MkDir caminhoExcel & Application.PathSeparator & "saida"
    On Error GoTo 0
    
    ' Salvar como texto
    Dim conteudoCompleto As String
    conteudoCompleto = Me.txtCabecalho.Value & Me.txtConteudo.Value & Me.txtRodape.Value
    
    ' Gravar arquivo
    Dim fileNum As Integer
    fileNum = FreeFile
    Open caminhoCompleto For Output As #fileNum
    Print #fileNum, conteudoCompleto
    Close #fileNum
    
    MsgBox "Relatório exportado com sucesso!" & vbCrLf & vbCrLf & _
           "Arquivo: " & nomeArquivo & vbCrLf & _
           "Local: " & caminhoExcel & "\saida\", _
           vbInformation, "Exportação Concluída"
End Sub

' ============================================
' Botão Fechar
' ============================================
Private Sub cmdFechar_Click()
    Unload Me
End Sub

' ============================================
' Inicialização do formulário
' ============================================
Private Sub UserForm_Initialize()
    ' Configurar tamanho aproximado A4 (595x842 pontos)
    Me.Width = 600
    Me.Height = 842
    
    ' Configurar fonte monoespaçada para melhor alinhamento
    Me.txtCabecalho.Font.Name = "Courier New"
    Me.txtCabecalho.Font.Size = 9
    
    Me.txtConteudo.Font.Name = "Courier New"
    Me.txtConteudo.Font.Size = 9
    
    Me.txtRodape.Font.Name = "Courier New"
    Me.txtRodape.Font.Size = 8
End Sub
