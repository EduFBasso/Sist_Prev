' ============================================
' modSimulacoes - Motor de cálculo
' ============================================

Option Explicit

Function CalcularTempo(ID_Cliente As Long) As Double
    Dim ws As Worksheet
    Dim ultima As Long
    Dim i As Long
    Dim n As Long
    Dim diasTotal As Double
    Dim dtInicio As Date
    Dim dtFim As Date
    Dim k As Long
    Dim tmp As Date

    Dim datasInicio() As Date
    Dim datasFim() As Date

    Set ws = Sheets("Vinculos")

    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To ultima
        If ws.Cells(i, 2).Value = ID_Cliente Then
            If IsDate(ws.Cells(i, 3).Value) Then
                dtInicio = CDate(ws.Cells(i, 3).Value)
            Else
                GoTo Proximo
            End If

            If IsDate(ws.Cells(i, 4).Value) Then
                dtFim = CDate(ws.Cells(i, 4).Value)
            Else
                dtFim = Date
            End If

            If dtFim >= dtInicio Then
                n = n + 1
                ReDim Preserve datasInicio(1 To n)
                ReDim Preserve datasFim(1 To n)
                datasInicio(n) = dtInicio
                datasFim(n) = dtFim
            End If
        End If
Proximo:
    Next i

    If n = 0 Then
        CalcularTempo = 0
        Exit Function
    End If

    Dim j As Long
    For i = 1 To n - 1
        For j = i + 1 To n
            If datasInicio(j) < datasInicio(i) Then
                tmp = datasInicio(i)
                datasInicio(i) = datasInicio(j)
                datasInicio(j) = tmp

                tmp = datasFim(i)
                datasFim(i) = datasFim(j)
                datasFim(j) = tmp
            End If
        Next j
    Next i

    Dim atualInicio As Date
    Dim atualFim As Date

    atualInicio = datasInicio(1)
    atualFim = datasFim(1)

    For k = 2 To n
        If datasInicio(k) <= atualFim Then
            If datasFim(k) > atualFim Then
                atualFim = datasFim(k)
            End If
        Else
            diasTotal = diasTotal + DateDiff("d", atualInicio, atualFim)
            atualInicio = datasInicio(k)
            atualFim = datasFim(k)
        End If
    Next k

    diasTotal = diasTotal + DateDiff("d", atualInicio, atualFim)

    CalcularTempo = diasTotal / 365.25
End Function

Function SimularAposentadoria(ID_Cliente As Long, regra As String) As Collection
    ' Usa GetParametro para aplicar regras
End Function

Function CalcularIdade(nascimento As Date) As Long
    CalcularIdade = DateDiff("yyyy", nascimento, Date)
    If Date < DateSerial(Year(Date), Month(nascimento), Day(nascimento)) Then
        CalcularIdade = CalcularIdade - 1
    End If
End Function

Function CalcularTempoEspecial(ID_Cliente As Long) As Double
    Dim ws As Worksheet
    Dim ultima As Long
    Dim i As Long
    Dim diasTotal As Double
    Dim dtInicio As Date
    Dim dtFim As Date
    Dim grau As String
    Dim fator As Double

    Set ws = Sheets("Vinculos")

    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To ultima
        If ws.Cells(i, 2).Value = ID_Cliente Then

            ' Coluna 6: Especial ("Sim" / "Não")
            If UCase(CStr(ws.Cells(i, 6).Value)) = "SIM" Then

                If IsDate(ws.Cells(i, 3).Value) Then
                    dtInicio = CDate(ws.Cells(i, 3).Value)
                Else
                    GoTo ProximoEspecial
                End If

                If IsDate(ws.Cells(i, 4).Value) Then
                    dtFim = CDate(ws.Cells(i, 4).Value)
                Else
                    dtFim = Date
                End If

                If dtFim >= dtInicio Then
                    grau = Trim(CStr(ws.Cells(i, 7).Value)) ' 15, 20 ou 25

                    Select Case grau
                        Case "15"
                            fator = Nz(GetParametro("Conversao_Especial_15"))
                        Case "20"
                            fator = Nz(GetParametro("Conversao_Especial_20"))
                        Case "25"
                            fator = Nz(GetParametro("Conversao_Especial_25"))
                        Case Else
                            fator = 1
                    End Select

                    If fator <= 0 Then fator = 1

                    diasTotal = diasTotal + DateDiff("d", dtInicio, dtFim) * fator
                End If
            End If
        End If
ProximoEspecial:
    Next i

    CalcularTempoEspecial = diasTotal / 365.25
End Function

Function RegraTempoContribuicao(ID_Cliente As Long) As Collection

    Dim resultado As New Collection
    Dim ws As Worksheet
    Dim linha As Long
    Dim sexo As String
    Dim tempoTotal As Double
    Dim tempoNecessario As Double
    Dim falta As Double
    Dim anosFaltantes As Long
    Dim mesesFaltantes As Long
    Dim diasFaltantes As Long
    Dim dataPrevista As Date
    Dim idadeProj As Long

    ' Buscar sexo do cliente
    Set ws = Sheets("Cadastro_Clientes")
    linha = BuscarLinhaCliente(ID_Cliente)
    sexo = ws.Cells(linha, 8).Value  ' Supondo que a coluna 6 é Sexo

    ' Tempo total já calculado
    tempoTotal = CalcularTempo(ID_Cliente)

    ' Define tempo necessário conforme sexo
    If UCase(sexo) = "M" Or UCase(sexo) = "MASCULINO" Then
        tempoNecessario = 35
    Else
        tempoNecessario = 30
    End If

    ' Calcula quanto falta
    falta = tempoNecessario - tempoTotal

    ' Se já tem direito
    If falta <= 0 Then
        resultado.Add "Sim", "Direito"
        resultado.Add tempoTotal, "TempoTotal"
        resultado.Add 0, "Falta"
        resultado.Add Date, "DataPrevista"
        resultado.Add CalcularIdade(ws.Cells(linha, 7).Value), "IdadeProjetada"
        resultado.Add "Direito adquirido pela regra antiga.", "Obs"
        Set RegraTempoContribuicao = resultado
        Exit Function
    End If

    ' Se não tem direito, calcular projeção
    ' Converte anos faltantes em dias
    Dim diasFalta As Long
    diasFalta = falta * 365.25

    dataPrevista = Date + diasFalta
    idadeProj = CalcularIdade(ws.Cells(linha, 7).Value) + falta

    resultado.Add "Não", "Direito"
    resultado.Add tempoTotal, "TempoTotal"
    resultado.Add falta, "Falta"
    resultado.Add dataPrevista, "DataPrevista"
    resultado.Add idadeProj, "IdadeProjetada"
    resultado.Add "Falta completar o tempo mínimo de contribuição.", "Obs"

    Set RegraTempoContribuicao = resultado

End Function

