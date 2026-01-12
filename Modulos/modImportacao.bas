' ============================================
' modImportacao - Importação de vínculos (CNIS)
' ============================================

Option Explicit

' ============================================
' IMPORTAÇÃO AUTOMÁTICA - Detecta e importa todos os CSVs relacionados
' ============================================

Sub ImportarCNIS_Automatico(Optional caminhoInicial As String = "")
    ' Importa automaticamente todos os arquivos CSV gerados pelo extrator Python
    ' Pode receber qualquer um dos arquivos CSV gerados (detecta o timestamp e importa todos)
    ' Se não receber parâmetro, pede ao usuário para selecionar um arquivo
    
    Dim caminho As String
    Dim pasta As String
    Dim prefixo As String
    Dim timestamp As String
    Dim ID_Cliente As Long
    Dim nomeCliente As String
    Dim wsCadastro As Worksheet
    Dim ultima As Long
    
    ' Se não foi passado arquivo, pedir para o usuário selecionar
    If caminhoInicial = "" Then
        Dim fd As FileDialog
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        
        With fd
            .Title = "Selecione um dos arquivos CSV gerados pelo extrator CNIS"
            .Filters.Clear
            .Filters.Add "Arquivos CSV", "*.csv"
            .AllowMultiSelect = False
            
            If .Show = -1 Then
                caminho = .SelectedItems(1)
            Else
                Exit Sub ' Usuário cancelou
            End If
        End With
    Else
        caminho = caminhoInicial
    End If
    
    ' Verificar se o arquivo existe
    If Len(Dir(caminho)) = 0 Then
        MsgBox "Arquivo não encontrado: " & caminho, vbExclamation
        Exit Sub
    End If
    
    ' Extrair pasta e detectar timestamp
    ' Exemplo: C:\pasta\saida\JOAO_CARLOS\extrato_202....csv
    '          → pasta = C:\pasta\saida\JOAO_CARLOS\
    '          → timestamp = 20260112_153946
    
    Dim nomeArquivo As String
    Dim posUltimaBarra As Long
    Dim separador As String
    
    ' Detectar separador (Windows \ ou macOS/Linux /)
    If InStr(caminho, "\\") > 0 Then
        separador = "\\"
        posUltimaBarra = InStrRev(caminho, "\\")
    Else
        separador = "/"
        posUltimaBarra = InStrRev(caminho, "/")
    End If
    
    pasta = Left(caminho, posUltimaBarra)
    nomeArquivo = Mid(caminho, posUltimaBarra + 1)
    
    ' Extrair timestamp do nome do arquivo
    ' Formato: extrato_AAAAMMDD_HHMMSS_tipo.csv
    ' Ou formato sem timestamp: saida_tipo.csv
    
    Dim partes() As String
    partes = Split(nomeArquivo, "_")
    
    If UBound(partes) >= 2 Then
        ' Tem timestamp: extrato_20260112_153946_dados_cliente.csv
        prefixo = partes(0) & "_"
        timestamp = partes(1) & "_" & partes(2)
    Else
        ' Sem timestamp: saida_dados_cliente.csv
        prefixo = partes(0) & "_"
        timestamp = ""
    End If
    
    ' Extrair nome da pasta do cliente (última pasta antes do arquivo)
    Dim posAntepenultimaBarra As Long
    Dim pastaCliente As String
    
    ' Remover separador final se existir
    Dim pastaSemSeparador As String
    pastaSemSeparador = Left(pasta, Len(pasta) - 1)
    
    ' Encontrar separador anterior para pegar nome da pasta
    If separador = "\\" Then
        posAntepenultimaBarra = InStrRev(pastaSemSeparador, "\\")
    Else
        posAntepenultimaBarra = InStrRev(pastaSemSeparador, "/")
    End If
    
    If posAntepenultimaBarra > 0 Then
        pastaCliente = Mid(pastaSemSeparador, posAntepenultimaBarra + 1)
    Else
        pastaCliente = "(pasta raiz)"
    End If
    
    ' Construir nomes dos arquivos
    Dim arquivoDados As String
    Dim arquivoVinculos As String
    Dim arquivoRemuneracoes As String
    
    If timestamp <> "" Then
        arquivoDados = pasta & prefixo & timestamp & "_dados_cliente.csv"
        arquivoVinculos = pasta & prefixo & timestamp & "_vinculos_estruturado.csv"
        arquivoRemuneracoes = pasta & prefixo & timestamp & "_remuneracoes.csv"
    Else
        arquivoDados = pasta & prefixo & "dados_cliente.csv"
        arquivoVinculos = pasta & prefixo & "vinculos_estruturado.csv"
        arquivoRemuneracoes = pasta & prefixo & "remuneracoes.csv"
    End If
    
    ' Informar pasta do cliente detectada
    Debug.Print "📁 Pasta do cliente: " & pastaCliente
    
    ' ========================================
    ' ETAPA 1: Importar dados do cliente
    ' ========================================
    
    If Len(Dir(arquivoDados)) = 0 Then
        MsgBox "Arquivo de dados do cliente não encontrado: " & vbCrLf & arquivoDados, vbExclamation
        Exit Sub
    End If
    
    ' Ler dados do CSV
    Dim fNum As Integer
    Dim linha As String
    Dim dadosCSV() As String
    
    fNum = FreeFile
    Open arquivoDados For Input As #fNum
    
    ' Pular cabeçalho
    Line Input #fNum, linha
    
    ' Ler dados (deve ter apenas uma linha)
    If Not EOF(fNum) Then
        Line Input #fNum, linha
        dadosCSV = Split(linha, ";")
    End If
    
    Close #fNum
    
    If UBound(dadosCSV) < 4 Then
        MsgBox "Arquivo de dados do cliente com formato inválido.", vbExclamation
        Exit Sub
    End If
    
    ' Dados do CSV: NIT;CPF;Nome;DataNascimento;NomeMae
    Dim nitCliente As String
    Dim cpfCliente As String
    Dim dataNascCliente As String
    Dim nomeMaeCliente As String
    
    nitCliente = Trim(dadosCSV(0))
    cpfCliente = Trim(dadosCSV(1))
    nomeCliente = Trim(dadosCSV(2))
    dataNascCliente = Trim(dadosCSV(3))
    nomeMaeCliente = Trim(dadosCSV(4))
    
    ' Verificar se cliente já existe (por CPF)
    Set wsCadastro = Sheets("Cadastro_Clientes")
    ultima = wsCadastro.Cells(wsCadastro.Rows.Count, 1).End(xlUp).Row
    
    ID_Cliente = 0
    Dim i As Long
    
    For i = 2 To ultima
        If Trim(CStr(wsCadastro.Cells(i, 3).Value)) = cpfCliente Then
            ID_Cliente = wsCadastro.Cells(i, 1).Value
            Exit For
        End If
    Next i
    
    ' Se não existir, criar novo cliente
    If ID_Cliente = 0 Then
        ID_Cliente = ultima  ' Novo ID
        
        wsCadastro.Cells(ultima + 1, 1).Value = ID_Cliente
        wsCadastro.Cells(ultima + 1, 2).Value = nomeCliente
        wsCadastro.Cells(ultima + 1, 3).Value = cpfCliente
        wsCadastro.Cells(ultima + 1, 4).Value = dataNascCliente
        wsCadastro.Cells(ultima + 1, 5).Value = nomeMaeCliente
        wsCadastro.Cells(ultima + 1, 6).Value = nitCliente
        
        MsgBox "Novo cliente criado: " & nomeCliente & " (ID " & ID_Cliente & ")", vbInformation
    Else
        MsgBox "Cliente já existe: " & nomeCliente & " (ID " & ID_Cliente & ")" & vbCrLf & _
               "Os vínculos e remunerações serão adicionados.", vbInformation
    End If
    
    ' ========================================
    ' ETAPA 2: Importar vínculos
    ' ========================================
    
    If Len(Dir(arquivoVinculos)) > 0 Then
        Call ImportarVinculosDeCSV(arquivoVinculos, ID_Cliente)
    Else
        MsgBox "Aviso: Arquivo de vínculos não encontrado: " & vbCrLf & arquivoVinculos, vbExclamation
    End If
    
    ' ========================================
    ' ETAPA 3: Importar remunerações
    ' ========================================
    
    If Len(Dir(arquivoRemuneracoes)) > 0 Then
        Call ImportarRemuneracoesDeCSV(arquivoRemuneracoes, ID_Cliente)
    Else
        MsgBox "Aviso: Arquivo de remunerações não encontrado: " & vbCrLf & arquivoRemuneracoes, vbExclamation
    End If
    
    ' ========================================
    ' CONCLUÍDO
    ' ========================================
    
    MsgBox "✅ IMPORTAÇÃO CONCLUÍDA!" & vbCrLf & vbCrLf & _
           "Cliente: " & nomeCliente & vbCrLf & _
           "ID: " & ID_Cliente & vbCrLf & _
           "CPF: " & cpfCliente & vbCrLf & vbCrLf & _
           "Use o formulário de Busca para localizar o cliente.", vbInformation, "Importação CNIS"
    
End Sub

Sub ImportarVinculosDeCSV(caminhoCSV As String, ID_Cliente As Long)

    Dim fNum As Integer
    Dim linha As String
    Dim partes() As String
    Dim primeiraLinha As Boolean
    Dim dados As Collection

    If Len(Dir(caminhoCSV)) = 0 Then
        MsgBox "Arquivo CSV não encontrado: " & caminhoCSV, vbExclamation
        Exit Sub
    End If

    fNum = FreeFile
    Open caminhoCSV For Input As #fNum

    primeiraLinha = True

    Do While Not EOF(fNum)
        Line Input #fNum, linha

        If primeiraLinha Then
            primeiraLinha = False
        Else
            If Trim(linha) <> "" Then
                partes = Split(linha, ";")

                If UBound(partes) >= 9 Then
                    Dim inicioStr As String
                    Dim fimStr As String
                    Dim ultRemStr As String
                    Dim dataFimCalculada As String

                    inicioStr = partes(7)
                    fimStr = partes(8)
                    ultRemStr = partes(9)

                    ' Regra: se não houver Data Fim, usar a última remuneração
                    ' como término (último dia da competência mm/aaaa)
                    If Trim(fimStr) = "" And Trim(ultRemStr) <> "" Then
                        Dim compParts() As String
                        compParts = Split(ultRemStr, "/")
                        If UBound(compParts) = 1 Then
                            Dim mes As Integer
                            Dim ano As Integer
                            On Error Resume Next
                            mes = CInt(compParts(0))
                            ano = CInt(compParts(1))
                            On Error GoTo 0
                            If mes >= 1 And mes <= 12 And ano > 0 Then
                                ' DateSerial com dia 0 do mês seguinte retorna
                                ' o último dia do mês informado
                                dataFimCalculada = Format$(DateSerial(ano, mes + 1, 0), "dd/mm/yyyy")
                                fimStr = dataFimCalculada
                            End If
                        End If
                    End If

                    Set dados = New Collection

                    dados.Add 0, "ID_Vinculo"
                    dados.Add ID_Cliente, "ID_Cliente"
                    dados.Add inicioStr, "Inicio"
                    dados.Add fimStr, "Fim"
                    dados.Add partes(6), "Tipo"
                    dados.Add "Não", "Especial"
                    dados.Add "", "Grau"
                    dados.Add False, "Rural"
                    dados.Add False, "Militar"
                    dados.Add False, "Exterior"
                    dados.Add False, "Concomitante"
                    dados.Add False, "Atraso"
                    dados.Add False, "Complementar"
                    dados.Add 0, "Salario"
                    dados.Add "Importado do CNIS - " & partes(5), "Observacoes"

                    Call SalvarVinculo(dados)
                End If
            End If
        End If
    Loop

    Close #fNum

    MsgBox "Importação de vínculos concluída.", vbInformation

End Sub

Sub ImportarRemuneracoesDeCSV(caminhoCSV As String, ID_Cliente As Long)
    ' Importa as remunerações do CSV gerado pelo script Python
    ' e associa ao vínculo correspondente do cliente
    
    Dim fNum As Integer
    Dim linha As String
    Dim partes() As String
    Dim primeiraLinha As Boolean
    Dim wsVinculos As Worksheet
    Dim wsRemuneracoes As Worksheet
    Dim ultimaV As Long
    Dim ultimaR As Long
    Dim i As Long
    Dim ID_Vinculo As Long
    Dim seq As String
    Dim codigoEmp As String
    Dim competencia As String
    Dim remuneracao As String
    Dim indicadores As String
    Dim linhaVinculo As Long
    Dim totalImportado As Long
    
    ' Verificar se o arquivo existe
    If Len(Dir(caminhoCSV)) = 0 Then
        MsgBox "Arquivo CSV não encontrado: " & caminhoCSV, vbExclamation
        Exit Sub
    End If
    
    Set wsVinculos = Sheets("Vinculos")
    
    ' Verificar se existe a planilha Remuneracoes, senão criar
    On Error Resume Next
    Set wsRemuneracoes = Sheets("Remuneracoes")
    On Error GoTo 0
    
    If wsRemuneracoes Is Nothing Then
        ' Criar planilha de remunerações
        Set wsRemuneracoes = Sheets.Add(After:=Sheets(Sheets.Count))
        wsRemuneracoes.Name = "Remuneracoes"
        
        ' Criar cabeçalho
        wsRemuneracoes.Cells(1, 1).Value = "ID_Remuneracao"
        wsRemuneracoes.Cells(1, 2).Value = "ID_Vinculo"
        wsRemuneracoes.Cells(1, 3).Value = "ID_Cliente"
        wsRemuneracoes.Cells(1, 4).Value = "Competencia"
        wsRemuneracoes.Cells(1, 5).Value = "Valor"
        wsRemuneracoes.Cells(1, 6).Value = "Indicadores"
        wsRemuneracoes.Cells(1, 7).Value = "Seq"
        wsRemuneracoes.Cells(1, 8).Value = "CodigoEmp"
        
        ' Formatar cabeçalho
        With wsRemuneracoes.Rows(1)
            .Font.Bold = True
            .Interior.Color = RGB(200, 220, 255)
        End With
    End If
    
    ' Abrir CSV
    fNum = FreeFile
    Open caminhoCSV For Input As #fNum
    
    primeiraLinha = True
    totalImportado = 0
    
    Do While Not EOF(fNum)
        Line Input #fNum, linha
        
        If primeiraLinha Then
            primeiraLinha = False
        Else
            If Trim(linha) <> "" Then
                partes = Split(linha, ";")
                
                ' CSV: Pagina, Seq, CodigoEmp, Competencia, Remuneracao, Indicadores
                If UBound(partes) >= 4 Then
                    seq = Trim(partes(1))
                    codigoEmp = Trim(partes(2))
                    competencia = Trim(partes(3))
                    remuneracao = Trim(partes(4))
                    
                    If UBound(partes) >= 5 Then
                        indicadores = Trim(partes(5))
                    Else
                        indicadores = ""
                    End If
                    
                    ' Buscar o ID_Vinculo correspondente a este código de empresa e cliente
                    ID_Vinculo = 0
                    ultimaV = wsVinculos.Cells(wsVinculos.Rows.Count, 1).End(xlUp).Row
                    
                    For i = 2 To ultimaV
                        ' Coluna 2 = ID_Cliente
                        If wsVinculos.Cells(i, 2).Value = ID_Cliente Then
                            ' Verificar se as observações contêm o código da empresa
                            Dim obs As String
                            obs = CStr(wsVinculos.Cells(i, 15).Value)
                            
                            If InStr(1, obs, codigoEmp, vbTextCompare) > 0 Then
                                ID_Vinculo = wsVinculos.Cells(i, 1).Value
                                Exit For
                            End If
                        End If
                    Next i
                    
                    ' Se encontrou o vínculo, adicionar a remuneração
                    If ID_Vinculo > 0 Then
                        ultimaR = wsRemuneracoes.Cells(wsRemuneracoes.Rows.Count, 1).End(xlUp).Row
                        
                        ' Verificar se já existe esta remuneração (evitar duplicatas)
                        Dim existe As Boolean
                        existe = False
                        
                        For i = 2 To ultimaR
                            If wsRemuneracoes.Cells(i, 2).Value = ID_Vinculo And _
                               wsRemuneracoes.Cells(i, 4).Value = competencia Then
                                existe = True
                                Exit For
                            End If
                        Next i
                        
                        If Not existe Then
                            Dim novaLinha As Long
                            novaLinha = ultimaR + 1
                            
                            ' Gerar novo ID
                            Dim novoID As Long
                            If ultimaR < 2 Then
                                novoID = 1
                            Else
                                novoID = wsRemuneracoes.Cells(ultimaR, 1).Value + 1
                            End If
                            
                            ' Salvar remuneração
                            wsRemuneracoes.Cells(novaLinha, 1).Value = novoID
                            wsRemuneracoes.Cells(novaLinha, 2).Value = ID_Vinculo
                            wsRemuneracoes.Cells(novaLinha, 3).Value = ID_Cliente
                            wsRemuneracoes.Cells(novaLinha, 4).Value = competencia
                            wsRemuneracoes.Cells(novaLinha, 5).Value = CDbl(Replace(remuneracao, ",", "."))
                            wsRemuneracoes.Cells(novaLinha, 6).Value = indicadores
                            wsRemuneracoes.Cells(novaLinha, 7).Value = seq
                            wsRemuneracoes.Cells(novaLinha, 8).Value = codigoEmp
                            
                            totalImportado = totalImportado + 1
                        End If
                    End If
                End If
            End If
        End If
    Loop
    
    Close #fNum
    
    MsgBox "Importação concluída: " & totalImportado & " remunerações importadas.", vbInformation
    
End Sub

Function CalcularMediaSalarios(ID_Cliente As Long) As Double
    ' Calcula a média dos 80% maiores salários do cliente
    ' baseado nas remunerações importadas do CNIS
    
    Dim ws As Worksheet
    Dim ultima As Long
    Dim i As Long
    Dim salarios() As Double
    Dim n As Long
    Dim total As Double
    Dim quantidadeConsiderar As Long
    
    On Error Resume Next
    Set ws = Sheets("Remuneracoes")
    On Error GoTo 0
    
    If ws Is Nothing Then
        CalcularMediaSalarios = 0
        Exit Function
    End If
    
    ' Coletar todos os salários do cliente
    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    n = 0
    
    For i = 2 To ultima
        If ws.Cells(i, 3).Value = ID_Cliente Then
            Dim valor As Double
            valor = ws.Cells(i, 5).Value
            
            If valor > 0 Then
                n = n + 1
                ReDim Preserve salarios(1 To n)
                salarios(n) = valor
            End If
        End If
    Next i
    
    If n = 0 Then
        CalcularMediaSalarios = 0
        Exit Function
    End If
    
    ' Ordenar salários em ordem decrescente (bubble sort simples)
    Dim j As Long
    Dim temp As Double
    
    For i = 1 To n - 1
        For j = i + 1 To n
            If salarios(j) > salarios(i) Then
                temp = salarios(i)
                salarios(i) = salarios(j)
                salarios(j) = temp
            End If
        Next j
    Next i
    
    ' Calcular média dos 80% maiores
    quantidadeConsiderar = Application.WorksheetFunction.Max(1, Int(n * 0.8))
    
    total = 0
    For i = 1 To quantidadeConsiderar
        total = total + salarios(i)
    Next i
    
    CalcularMediaSalarios = total / quantidadeConsiderar
    
End Function

Sub CriarRelatorioCliente(ID_Cliente As Long)
    ' Cria uma planilha resumo com todos os dados organizados do cliente
    ' Baseado nos dados extraídos do CNIS
    
    Dim wsCadastro As Worksheet
    Dim wsVinculos As Worksheet
    Dim wsRelatorio As Worksheet
    Dim nomeCliente As String
    Dim cpfCliente As String
    Dim i As Long
    Dim linha As Long
    Dim nomeAba As String
    
    ' Busca dados do cliente
    Set wsCadastro = Sheets("Cadastro_Clientes")
    
    For i = 2 To wsCadastro.Cells(wsCadastro.Rows.Count, 1).End(xlUp).Row
        If wsCadastro.Cells(i, 1).Value = ID_Cliente Then
            nomeCliente = wsCadastro.Cells(i, 2).Value
            cpfCliente = wsCadastro.Cells(i, 3).Value
            Exit For
        End If
    Next i
    
    If nomeCliente = "" Then
        MsgBox "Cliente não encontrado (ID " & ID_Cliente & ")", vbExclamation
        Exit Sub
    End If
    
    ' Nome da aba: Primeiros 25 caracteres do nome (limite do Excel: 31)
    nomeAba = "Rel_" & Left(nomeCliente, 25)
    
    ' Remove aba se já existir
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets(nomeAba).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Cria nova aba
    Set wsRelatorio = Sheets.Add(After:=Sheets(Sheets.Count))
    wsRelatorio.Name = nomeAba
    
    ' CABEÇALHO DO RELATÓRIO
    linha = 1
    With wsRelatorio
        .Cells(linha, 1).Value = "RELATÓRIO DE VÍNCULOS E CONTRIBUIÇÕES - CNIS"
        .Cells(linha, 1).Font.Bold = True
        .Cells(linha, 1).Font.Size = 14
        linha = linha + 1
        
        linha = linha + 1
        .Cells(linha, 1).Value = "Cliente:"
        .Cells(linha, 1).Font.Bold = True
        .Cells(linha, 2).Value = nomeCliente
        linha = linha + 1
        
        .Cells(linha, 1).Value = "CPF:"
        .Cells(linha, 1).Font.Bold = True
        .Cells(linha, 2).Value = cpfCliente
        linha = linha + 1
        
        .Cells(linha, 1).Value = "Data:"
        .Cells(linha, 1).Font.Bold = True
        .Cells(linha, 2).Value = Date
        linha = linha + 2
        
        ' VÍNCULOS
        .Cells(linha, 1).Value = "VÍNCULOS CONTRIBUTIVOS"
        .Cells(linha, 1).Font.Bold = True
        .Cells(linha, 1).Font.Size = 12
        linha = linha + 1
        
        ' Cabeçalho da tabela
        .Cells(linha, 1).Value = "Seq"
        .Cells(linha, 2).Value = "Tipo"
        .Cells(linha, 3).Value = "Empresa/CNPJ"
        .Cells(linha, 4).Value = "Data Início"
        .Cells(linha, 5).Value = "Data Fim"
        .Cells(linha, 6).Value = "Tempo (anos)"
        .Cells(linha, 7).Value = "Indicadores"
        
        ' Formatar cabeçalho
        .Range(.Cells(linha, 1), .Cells(linha, 7)).Font.Bold = True
        .Range(.Cells(linha, 1), .Cells(linha, 7)).Interior.Color = RGB(200, 200, 200)
        linha = linha + 1
    End With
    
    ' Busca vínculos do cliente
    Set wsVinculos = Sheets("Vinculos")
    Dim linhaInicio As Long
    linhaInicio = linha
    
    For i = 2 To wsVinculos.Cells(wsVinculos.Rows.Count, 1).End(xlUp).Row
        If wsVinculos.Cells(i, 3).Value = ID_Cliente Then
            Dim seq As String
            Dim tipoFiliado As String
            Dim empresaCNPJ As String
            Dim dataIni As String
            Dim dataFim As String
            Dim tempoAnos As Double
            Dim indicadores As String
            
            seq = wsVinculos.Cells(i, 4).Value  ' Seq
            tipoFiliado = wsVinculos.Cells(i, 7).Value  ' Tipo
            
            ' Monta campo Empresa/CNPJ
            Dim cnpj As String
            Dim empresa As String
            cnpj = wsVinculos.Cells(i, 5).Value
            empresa = wsVinculos.Cells(i, 6).Value
            
            If Len(cnpj) > 0 Then
                empresaCNPJ = empresa & " (" & cnpj & ")"
            Else
                empresaCNPJ = tipoFiliado
            End If
            
            dataIni = wsVinculos.Cells(i, 8).Value
            dataFim = wsVinculos.Cells(i, 9).Value
            indicadores = wsVinculos.Cells(i, 11).Value
            
            ' Calcula tempo
            If IsDate(dataIni) Then
                Dim dtFim As Date
                If IsDate(dataFim) Then
                    dtFim = CDate(dataFim)
                Else
                    dtFim = Date
                End If
                tempoAnos = (dtFim - CDate(dataIni)) / 365.25
            End If
            
            ' Preenche linha
            With wsRelatorio
                .Cells(linha, 1).Value = seq
                .Cells(linha, 2).Value = tipoFiliado
                .Cells(linha, 3).Value = empresaCNPJ
                .Cells(linha, 4).Value = dataIni
                .Cells(linha, 5).Value = dataFim
                .Cells(linha, 6).Value = Round(tempoAnos, 2)
                .Cells(linha, 7).Value = indicadores
            End With
            
            linha = linha + 1
        End If
    Next i
    
    ' Formata tabela de vínculos
    If linha > linhaInicio Then
        With wsRelatorio
            .Range(.Cells(linhaInicio, 1), .Cells(linha - 1, 7)).Borders.LineStyle = xlContinuous
        End With
    End If
    
    ' Ajusta largura das colunas
    wsRelatorio.Columns("A:G").AutoFit
    
    ' Vai para a aba criada
    wsRelatorio.Activate
    wsRelatorio.Range("A1").Select
    
    MsgBox "Relatório criado com sucesso: " & nomeAba, vbInformation
    
End Sub