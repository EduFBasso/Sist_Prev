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
