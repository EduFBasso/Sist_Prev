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
                    Set dados = New Collection

                    dados.Add 0, "ID_Vinculo"
                    dados.Add ID_Cliente, "ID_Cliente"
                    dados.Add partes(7), "Inicio"
                    dados.Add partes(8), "Fim"
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
