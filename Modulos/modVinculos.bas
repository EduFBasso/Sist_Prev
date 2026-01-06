' ============================================
' modVinculos - Operações dos vínculos
' ============================================

Option Explicit

Sub SalvarVinculo(dados As Collection)
    Dim ws As Worksheet
    Dim linha As Long
    Dim idV As Long
    
    Set ws = Sheets("Vinculos")
    
    If dados("ID_Vinculo") = 0 Then
        linha = ProximaLinha(ws)
        idV = NovoID("Vinculos", 1)
    Else
        linha = BuscarLinhaVinculo(dados("ID_Vinculo"))
        idV = dados("ID_Vinculo")
    End If
    
    ws.Cells(linha, 1).Value = idV
    ws.Cells(linha, 2).Value = dados("ID_Cliente")
    ws.Cells(linha, 3).Value = dados("Inicio")
    ws.Cells(linha, 4).Value = dados("Fim")
    ws.Cells(linha, 5).Value = dados("Tipo")
    ws.Cells(linha, 6).Value = dados("Especial")
    ws.Cells(linha, 7).Value = dados("Grau")
    ws.Cells(linha, 8).Value = dados("Rural")
    ws.Cells(linha, 9).Value = dados("Militar")
    ws.Cells(linha, 10).Value = dados("Exterior")
    ws.Cells(linha, 11).Value = dados("Concomitante")
    ws.Cells(linha, 12).Value = dados("Atraso")
    ws.Cells(linha, 13).Value = dados("Complementar")
    ws.Cells(linha, 14).Value = dados("Salario")
    ws.Cells(linha, 15).Value = dados("Observacoes")
    
    Call AtualizarIndicadoresCliente(dados("ID_Cliente"))
End Sub


Sub ExcluirVinculo(ID_Vinculo As Long)

    Dim ws As Worksheet
    Dim linha As Long
    Dim ID_Cliente As Long

    Set ws = Sheets("Vinculos")
    linha = BuscarLinhaVinculo(ID_Vinculo)

    If linha = 0 Then
        MsgBox "Vínculo não encontrado.", vbExclamation
        Exit Sub
    End If

    ID_Cliente = ws.Cells(linha, 2).Value

    ws.Rows(linha).ClearContents

    Call AtualizarIndicadoresCliente(ID_Cliente)

End Sub


Function BuscarLinhaVinculo(ID_Vinculo As Long) As Long

    Dim ws As Worksheet
    Dim ultima As Long
    Dim i As Long

    Set ws = Sheets("Vinculos")
    ultima = ws.Cells(ws.Rows.count, 1).End(xlUp).Row

    For i = 2 To ultima
        If ws.Cells(i, 1).Value = ID_Vinculo Then
            BuscarLinhaVinculo = i
            Exit Function
        End If
    Next i

    BuscarLinhaVinculo = 0

End Function

Sub CarregarVinculosCliente(ID_Cliente As Long)

    Dim ws As Worksheet
    Dim ultima As Long
    Dim i As Long

    Set ws = Sheets("Vinculos")

    frmCadastro.lstVinculos.Clear

    ultima = ws.Cells(ws.Rows.count, 1).End(xlUp).Row

    For i = 2 To ultima
        If ws.Cells(i, 2).Value = ID_Cliente Then

            frmCadastro.lstVinculos.AddItem ws.Cells(i, 1).Value ' ID_Vinculo
            frmCadastro.lstVinculos.List(frmCadastro.lstVinculos.ListCount - 1, 1) = ws.Cells(i, 3).Value ' Inicio
            frmCadastro.lstVinculos.List(frmCadastro.lstVinculos.ListCount - 1, 2) = ws.Cells(i, 4).Value ' Fim
            frmCadastro.lstVinculos.List(frmCadastro.lstVinculos.ListCount - 1, 3) = ws.Cells(i, 5).Value ' Tipo
            frmCadastro.lstVinculos.List(frmCadastro.lstVinculos.ListCount - 1, 4) = ws.Cells(i, 6).Value ' Especial
            frmCadastro.lstVinculos.List(frmCadastro.lstVinculos.ListCount - 1, 5) = ws.Cells(i, 7).Value ' Grau
            frmCadastro.lstVinculos.List(frmCadastro.lstVinculos.ListCount - 1, 6) = ws.Cells(i, 8).Value ' Salario

        End If
    Next i

End Sub


Sub CarregarVinculo(ID_Vinculo As Long)

    Dim ws As Worksheet
    Dim linha As Long

    Set ws = Sheets("Vinculos")
    linha = BuscarLinhaVinculo(ID_Vinculo)

    If linha = 0 Then
        MsgBox "Vínculo não encontrado.", vbExclamation
        Exit Sub
    End If

    With frmVinculos
        .txtIDVinculo.Value = ws.Cells(linha, 1).Value
        .txtIDCliente.Value = ws.Cells(linha, 2).Value
        .txtInicio.Value = ws.Cells(linha, 3).Value
        .txtFim.Value = ws.Cells(linha, 4).Value
        .cboTipo.Value = ws.Cells(linha, 5).Value
        .cboEspecial.Value = ws.Cells(linha, 6).Value
        .cboGrau.Value = ws.Cells(linha, 7).Value
        
        .chkRural.Value = ws.Cells(linha, 8).Value
        .chkMilitar.Value = ws.Cells(linha, 9).Value
        .chkExterior.Value = ws.Cells(linha, 10).Value
        .chkConcomitante.Value = ws.Cells(linha, 11).Value
        .chkAtraso.Value = ws.Cells(linha, 12).Value
        .chkComplementar.Value = ws.Cells(linha, 13).Value
        
        .txtSalario.Value = ws.Cells(linha, 14).Value
        .txtObs.Value = ws.Cells(linha, 15).Value
    End With

End Sub

Sub AtualizarIndicadoresCliente(ID_Cliente As Long)

    Dim wsV As Worksheet
    Dim wsC As Worksheet
    Dim ultima As Long
    Dim i As Long
    Dim linhaCliente As Long

    Dim temEspecial As Boolean
    Dim temRural As Boolean
    Dim temMilitar As Boolean
    Dim temExterior As Boolean
    Dim temConcomitante As Boolean
    Dim temAtraso As Boolean
    Dim temComplementar As Boolean

    Set wsV = Sheets("Vinculos")
    Set wsC = Sheets("Cadastro_Clientes")

    ' Localiza o cliente na planilha de cadastro
    linhaCliente = BuscarLinhaCliente(ID_Cliente)
    If linhaCliente = 0 Then Exit Sub

    ' Varre todos os vínculos do cliente
    ultima = wsV.Cells(wsV.Rows.count, 1).End(xlUp).Row

    For i = 2 To ultima
        If wsV.Cells(i, 2).Value = ID_Cliente Then

            If wsV.Cells(i, 6).Value = "Sim" Then temEspecial = True
            If wsV.Cells(i, 8).Value = True Then temRural = True
            If wsV.Cells(i, 9).Value = True Then temMilitar = True
            If wsV.Cells(i, 10).Value = True Then temExterior = True
            If wsV.Cells(i, 11).Value = True Then temConcomitante = True
            If wsV.Cells(i, 12).Value = True Then temAtraso = True
            If wsV.Cells(i, 13).Value = True Then temComplementar = True

        End If
    Next i

    ' Atualiza o cadastro do cliente
    wsC.Cells(linhaCliente, 22).Value = temEspecial
    wsC.Cells(linhaCliente, 23).Value = temRural
    wsC.Cells(linhaCliente, 24).Value = temMilitar
    wsC.Cells(linhaCliente, 25).Value = temExterior
    wsC.Cells(linhaCliente, 26).Value = temConcomitante
    wsC.Cells(linhaCliente, 27).Value = temAtraso
    wsC.Cells(linhaCliente, 28).Value = temComplementar

End Sub

