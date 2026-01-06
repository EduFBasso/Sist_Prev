Attribute VB_Name = "ModPrincipal"
Option Explicit

' Módulo Principal - Sistema de Planejamento Previdenciário RGPS
' Inicialização e funções principais

' Inicializar o sistema completo
Public Sub InicializarSistema()
    On Error GoTo ErroHandler
    
    Application.ScreenUpdating = False
    
    ' Inicializar todas as planilhas necessárias
    ModClientes.InicializarPlanilhaClientes
    ModImportacaoINSS.InicializarPlanilhaVinculos
    ModImportacaoINSS.InicializarPlanilhaCalculoTempo
    ModSimulacaoAposentadoria.InicializarPlanilhaResultados
    InicializarPlanilhaPrincipal
    
    Application.ScreenUpdating = True
    
    MsgBox "Sistema inicializado com sucesso!" & vbCrLf & vbCrLf & _
           "Planilhas criadas:" & vbCrLf & _
           "- Início (Dashboard)" & vbCrLf & _
           "- Clientes" & vbCrLf & _
           "- Vínculos" & vbCrLf & _
           "- Cálculo de Tempo" & vbCrLf & _
           "- Resultados de Simulação", vbInformation, "Sistema Pronto"
    
    Exit Sub
    
ErroHandler:
    Application.ScreenUpdating = True
    MsgBox "Erro ao inicializar sistema: " & Err.Description, vbCritical, "Erro"
End Sub

' Criar planilha principal/dashboard
Private Sub InicializarPlanilhaPrincipal()
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Inicio")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Inicio"
        ws.Move before:=ThisWorkbook.Worksheets(1)
    End If
    
    ' Limpar planilha
    ws.Cells.Clear
    
    ' Criar cabeçalho
    With ws.Range("A1:F1")
        .Merge
        .Value = "SISTEMA DE PLANEJAMENTO PREVIDENCIÁRIO - RGPS"
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 40
    End With
    
    ' Informações do sistema
    ws.Range("A3").Value = "Bem-vindo ao Sistema de Planejamento Previdenciário!"
    ws.Range("A3").Font.Size = 12
    ws.Range("A3").Font.Bold = True
    
    ws.Range("A5").Value = "Este sistema permite:"
    ws.Range("A6").Value = "• Cadastrar clientes"
    ws.Range("A7").Value = "• Importar vínculos do INSS"
    ws.Range("A8").Value = "• Calcular tempos de contribuição"
    ws.Range("A9").Value = "• Simular aposentadorias (idade, tempo, pontos, especial)"
    ws.Range("A10").Value = "• Gerar relatórios para análise previdenciária"
    
    ' Menu rápido
    ws.Range("A12").Value = "MENU RÁPIDO DE FUNÇÕES:"
    ws.Range("A12").Font.Bold = True
    ws.Range("A12").Font.Size = 11
    
    ws.Range("A14").Value = "1. Cadastrar Novo Cliente"
    ws.Range("A15").Value = "2. Importar Vínculos"
    ws.Range("A16").Value = "3. Simular Aposentadorias"
    ws.Range("A17").Value = "4. Ver Clientes Cadastrados"
    
    ' Instruções
    ws.Range("A19").Value = "INSTRUÇÕES DE USO:"
    ws.Range("A19").Font.Bold = True
    ws.Range("A19").Font.Size = 11
    
    ws.Range("A21").Value = "1. Execute a macro 'InicializarSistema' para criar as planilhas (apenas na primeira vez)"
    ws.Range("A22").Value = "2. Use 'CadastrarCliente' para adicionar novos clientes na planilha Clientes"
    ws.Range("A23").Value = "3. Use 'AdicionarVinculo' ou 'ImportarVinculosCSV' para registrar vínculos empregatícios"
    ws.Range("A24").Value = "4. Use 'SimularAposentadorias' para calcular elegibilidade para diferentes tipos de aposentadoria"
    ws.Range("A25").Value = "5. Consulte a planilha 'ResultadosSimulacao' para ver os resultados detalhados"
    
    ' Formato CSV para importação
    ws.Range("A27").Value = "FORMATO DE ARQUIVO CSV PARA IMPORTAÇÃO:"
    ws.Range("A27").Font.Bold = True
    ws.Range("A27").Font.Size = 11
    
    ws.Range("A29").Value = "O arquivo CSV deve ter as seguintes colunas separadas por ponto-e-vírgula (;):"
    ws.Range("A30").Value = "Empresa;CNPJ;Data Início;Data Fim;Tipo Vínculo;Condição Especial"
    ws.Range("A30").Font.Italic = True
    
    ws.Range("A32").Value = "Exemplo:"
    ws.Range("A33").Value = "Empresa ABC Ltda;12.345.678/0001-90;01/01/2010;31/12/2015;CLT;Normal"
    ws.Range("A34").Value = "Empresa XYZ SA;98.765.432/0001-10;01/01/2016;;CLT;Insalubre"
    ws.Range("A33:A34").Font.Italic = True
    
    ' Informações de rodapé
    ws.Range("A36").Value = "Versão: 1.0 | Desenvolvido para planejamento previdenciário RGPS"
    ws.Range("A36").Font.Size = 9
    ws.Range("A36").Font.Italic = True
    
    ' Ajustar largura das colunas
    ws.Columns("A:F").AutoFit
    ws.Columns("A").ColumnWidth = 100
End Sub

' Exemplo de uso: Cadastrar cliente via código
Public Sub ExemploCadastrarCliente()
    Dim sucesso As Boolean
    
    sucesso = ModClientes.CadastrarCliente( _
        nome:="João da Silva", _
        cpf:="123.456.789-00", _
        dataNascimento:=DateSerial(1965, 5, 15), _
        sexo:="M", _
        email:="joao.silva@email.com", _
        telefone:="(11) 98765-4321" _
    )
    
    If sucesso Then
        Debug.Print "Cliente cadastrado com sucesso!"
    End If
End Sub

' Exemplo de uso: Adicionar vínculo
Public Sub ExemploAdicionarVinculo()
    Dim sucesso As Boolean
    
    sucesso = ModImportacaoINSS.AdicionarVinculo( _
        cpfCliente:="123.456.789-00", _
        empresa:="Empresa ABC Ltda", _
        cnpj:="12.345.678/0001-90", _
        dataInicio:=DateSerial(2000, 1, 1), _
        dataFim:=DateSerial(2020, 12, 31), _
        tipoVinculo:="CLT", _
        condicaoEspecial:="Normal" _
    )
    
    If sucesso Then
        ModImportacaoINSS.CalcularTemposContribuicao "123.456.789-00"
        Debug.Print "Vínculo adicionado e tempo calculado!"
    End If
End Sub

' Exemplo de uso: Simular aposentadorias
Public Sub ExemploSimularAposentadoria()
    Dim resultado As String
    
    resultado = ModSimulacaoAposentadoria.SimularAposentadorias("123.456.789-00")
    
    MsgBox resultado, vbInformation, "Resultados da Simulação"
End Sub

' Função para demonstração completa
Public Sub DemonstracaoCompleta()
    ' Inicializar sistema
    InicializarSistema
    
    ' Cadastrar cliente de exemplo
    Call ExemploCadastrarCliente
    
    ' Adicionar vínculo de exemplo
    Call ExemploAdicionarVinculo
    
    ' Simular aposentadoria
    Call ExemploSimularAposentadoria
    
    MsgBox "Demonstração completa executada! Verifique as planilhas.", vbInformation, "Concluído"
End Sub
