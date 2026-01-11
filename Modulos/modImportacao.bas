' ============================================
' modImportacao - Importação de vínculos (CNIS)
' ============================================

Option Explicit

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
