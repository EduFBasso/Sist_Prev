Attribute VB_Name = "ModImportacaoINSS"
Option Explicit

' Módulo de Importação de Extratos do INSS
' Sistema de Planejamento Previdenciário RGPS

' Função para importar vínculos de arquivo CSV
Public Function ImportarVinculosCSV(cpfCliente As String, caminhoArquivo As String) As Boolean
    On Error GoTo ErroHandler
    
    Dim ws As Worksheet
    Dim linha As String
    Dim dados() As String
    Dim fileNum As Integer
    Dim linhaAtual As Long
    Dim idCliente As Long
    
    ' Verificar se cliente existe
    idCliente = ModClientes.BuscarClientePorCPF(cpfCliente)
    If idCliente = 0 Then
        MsgBox "Cliente não encontrado. Cadastre o cliente primeiro.", vbExclamation, "Erro"
        ImportarVinculosCSV = False
        Exit Function
    End If
    
    ' Referência à planilha de vínculos
    Set ws = ThisWorkbook.Worksheets("Vinculos")
    linhaAtual = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Abrir arquivo CSV
    fileNum = FreeFile
    Open caminhoArquivo For Input As #fileNum
    
    ' Pular cabeçalho se existir
    If Not EOF(fileNum) Then
        Line Input #fileNum, linha
    End If
    
    ' Ler dados
    Do While Not EOF(fileNum)
        Line Input #fileNum, linha
        dados = Split(linha, ";")
        
        If UBound(dados) >= 5 Then
            ' Importar dados do vínculo
            ws.Cells(linhaAtual, 1).Value = linhaAtual - 1 ' ID
            ws.Cells(linhaAtual, 2).Value = cpfCliente
            ws.Cells(linhaAtual, 3).Value = dados(0) ' Empresa/Empregador
            ws.Cells(linhaAtual, 4).Value = dados(1) ' CNPJ
            ws.Cells(linhaAtual, 5).Value = CDate(dados(2)) ' Data Início
            ws.Cells(linhaAtual, 6).Value = IIf(dados(3) <> "", CDate(dados(3)), "") ' Data Fim
            ws.Cells(linhaAtual, 7).Value = dados(4) ' Tipo Vínculo
            ws.Cells(linhaAtual, 8).Value = IIf(UBound(dados) >= 5, dados(5), "Normal") ' Condição Especial
            
            linhaAtual = linhaAtual + 1
        End If
    Loop
    
    Close #fileNum
    
    ' Calcular tempos de contribuição
    CalcularTemposContribuicao cpfCliente
    
    MsgBox "Vínculos importados com sucesso!", vbInformation, "Sucesso"
    ImportarVinculosCSV = True
    Exit Function
    
ErroHandler:
    If fileNum <> 0 Then Close #fileNum
    MsgBox "Erro ao importar vínculos: " & Err.Description, vbCritical, "Erro"
    ImportarVinculosCSV = False
End Function

' Função para adicionar vínculo manualmente
Public Function AdicionarVinculo(cpfCliente As String, empresa As String, cnpj As String, _
                                dataInicio As Date, dataFim As Variant, tipoVinculo As String, _
                                condicaoEspecial As String) As Boolean
    On Error GoTo ErroHandler
    
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim idCliente As Long
    
    ' Verificar se cliente existe
    idCliente = ModClientes.BuscarClientePorCPF(cpfCliente)
    If idCliente = 0 Then
        MsgBox "Cliente não encontrado.", vbExclamation, "Erro"
        AdicionarVinculo = False
        Exit Function
    End If
    
    Set ws = ThisWorkbook.Worksheets("Vinculos")
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Inserir dados do vínculo
    ws.Cells(ultimaLinha, 1).Value = ultimaLinha - 1
    ws.Cells(ultimaLinha, 2).Value = cpfCliente
    ws.Cells(ultimaLinha, 3).Value = empresa
    ws.Cells(ultimaLinha, 4).Value = cnpj
    ws.Cells(ultimaLinha, 5).Value = dataInicio
    ws.Cells(ultimaLinha, 6).Value = IIf(IsDate(dataFim), dataFim, "")
    ws.Cells(ultimaLinha, 7).Value = tipoVinculo
    ws.Cells(ultimaLinha, 8).Value = condicaoEspecial
    
    MsgBox "Vínculo adicionado com sucesso!", vbInformation, "Sucesso"
    AdicionarVinculo = True
    Exit Function
    
ErroHandler:
    MsgBox "Erro ao adicionar vínculo: " & Err.Description, vbCritical, "Erro"
    AdicionarVinculo = False
End Function

' Função para calcular tempos de contribuição
Public Sub CalcularTemposContribuicao(cpfCliente As String)
    Dim wsVinculos As Worksheet
    Dim wsCalculo As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim tempoTotal As Double
    Dim tempoEspecial As Double
    Dim dataIni As Date
    Dim dataFim As Date
    Dim diasContribuicao As Long
    
    Set wsVinculos = ThisWorkbook.Worksheets("Vinculos")
    Set wsCalculo = ThisWorkbook.Worksheets("CalculoTempo")
    
    ultimaLinha = wsVinculos.Cells(wsVinculos.Rows.Count, 2).End(xlUp).Row
    
    tempoTotal = 0
    tempoEspecial = 0
    
    ' Percorrer vínculos do cliente
    For i = 2 To ultimaLinha
        If wsVinculos.Cells(i, 2).Value = cpfCliente Then
            dataIni = wsVinculos.Cells(i, 5).Value
            
            If wsVinculos.Cells(i, 6).Value <> "" Then
                dataFim = wsVinculos.Cells(i, 6).Value
            Else
                dataFim = Date ' Data atual se vínculo ativo
            End If
            
            diasContribuicao = dataFim - dataIni
            tempoTotal = tempoTotal + diasContribuicao
            
            ' Verificar se é tempo especial
            If wsVinculos.Cells(i, 8).Value <> "Normal" And wsVinculos.Cells(i, 8).Value <> "" Then
                tempoEspecial = tempoEspecial + diasContribuicao
            End If
        End If
    Next i
    
    ' Gravar resultados
    Dim linhaResultado As Long
    linhaResultado = wsCalculo.Cells(wsCalculo.Rows.Count, 1).End(xlUp).Row + 1
    
    wsCalculo.Cells(linhaResultado, 1).Value = cpfCliente
    wsCalculo.Cells(linhaResultado, 2).Value = tempoTotal / 365 ' Anos
    wsCalculo.Cells(linhaResultado, 3).Value = Int((tempoTotal Mod 365) / 30) ' Meses
    wsCalculo.Cells(linhaResultado, 4).Value = (tempoTotal Mod 365) Mod 30 ' Dias
    wsCalculo.Cells(linhaResultado, 5).Value = tempoEspecial / 365 ' Anos especiais
    wsCalculo.Cells(linhaResultado, 6).Value = Now() ' Data do cálculo
End Sub

' Inicializar planilha de vínculos
Public Sub InicializarPlanilhaVinculos()
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Vinculos")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Vinculos"
    End If
    
    ' Criar cabeçalhos se não existirem
    If ws.Cells(1, 1).Value = "" Then
        ws.Cells(1, 1).Value = "ID"
        ws.Cells(1, 2).Value = "CPF Cliente"
        ws.Cells(1, 3).Value = "Empresa"
        ws.Cells(1, 4).Value = "CNPJ"
        ws.Cells(1, 5).Value = "Data Início"
        ws.Cells(1, 6).Value = "Data Fim"
        ws.Cells(1, 7).Value = "Tipo Vínculo"
        ws.Cells(1, 8).Value = "Condição Especial"
        
        ' Formatar cabeçalhos
        With ws.Range("A1:H1")
            .Font.Bold = True
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With
        
        ws.Columns("A:H").AutoFit
    End If
End Sub

' Inicializar planilha de cálculo de tempo
Public Sub InicializarPlanilhaCalculoTempo()
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("CalculoTempo")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "CalculoTempo"
    End If
    
    ' Criar cabeçalhos se não existirem
    If ws.Cells(1, 1).Value = "" Then
        ws.Cells(1, 1).Value = "CPF Cliente"
        ws.Cells(1, 2).Value = "Anos Contribuição"
        ws.Cells(1, 3).Value = "Meses"
        ws.Cells(1, 4).Value = "Dias"
        ws.Cells(1, 5).Value = "Anos Especiais"
        ws.Cells(1, 6).Value = "Data Cálculo"
        
        ' Formatar cabeçalhos
        With ws.Range("A1:F1")
            .Font.Bold = True
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With
        
        ws.Columns("A:F").AutoFit
    End If
End Sub
