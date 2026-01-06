' ============================================
' modSimulacoes - Motor de cálculo
' ============================================

Option Explicit

Function CalcularTempo(ID_Cliente As Long) As Double
    ' Soma períodos da aba Vinculos
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
    ' Por enquanto retorna 0
    CalcularTempoEspecial = 0
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

