' ============================================
' modVinculos - Operações dos vínculos
' ============================================

Option Explicit

Sub SalvarVinculo(dados As Collection)
    Dim ws As Worksheet
    Dim linha As Long
    Dim idV As Long
    Dim dataInicio As Variant
    Dim dataFim As Variant
    
    Set ws = Sheets("Vinculos")
    
    If dados("ID_Vinculo") = 0 Then
        linha = ProximaLinha(ws)
        idV = NovoID("Vinculos", 1)
    Else
        linha = BuscarLinhaVinculo(dados("ID_Vinculo"))
        idV = dados("ID_Vinculo")
    End If
    
    ' Converte datas do formato brasileiro (DD/MM/YYYY) corretamente
    dataInicio = ConverteBRParaData(dados("Inicio"))
    dataFim = ConverteBRParaData(dados("Fim"))
    
    ' Ordem correta: ID_Vinculo, ID_Cliente, Data_Inicio, Data_Fim, Tipo_Vinculo,
    '                Especial, Grau_Especial, Salario_Contribuicao, Observacoes, Data_Cadastro,
    '                Rural, Militar, Exterior, Concomitante, Atraso, Complementar, Codigo Emp.
    
    ws.Cells(linha, 1).Value = idV
    ws.Cells(linha, 2).Value = dados("ID_Cliente")
    
    ' Formatar datas como texto para evitar auto-conversão do Excel
    ws.Cells(linha, 3).NumberFormat = "@"
    ws.Cells(linha, 3).Value = CStr(dados("Inicio"))
    
    ws.Cells(linha, 4).NumberFormat = "@"
    ws.Cells(linha, 4).Value = CStr(dados("Fim"))
    
    ws.Cells(linha, 5).Value = dados("Tipo")
    ws.Cells(linha, 6).Value = dados("Especial")
    ws.Cells(linha, 7).Value = dados("Grau")
    ws.Cells(linha, 8).Value = dados("Salario")
    
    ' Garantir que Observações seja salvo como texto (coluna 9)
    ws.Cells(linha, 9).NumberFormat = "@"
    ws.Cells(linha, 9).Value = CStr(dados("Observacoes"))
    
    ws.Cells(linha, 10).Value = Date  ' Data_Cadastro
    
    ' Garantir booleanos explícitos (não converter para texto em português)
    ws.Cells(linha, 11).Value = CBool(dados("Rural"))
    ws.Cells(linha, 12).Value = CBool(dados("Militar"))
    ws.Cells(linha, 13).Value = CBool(dados("Exterior"))
    ws.Cells(linha, 14).Value = CBool(dados("Concomitante"))
    ws.Cells(linha, 15).Value = CBool(dados("Atraso"))
    ws.Cells(linha, 16).Value = CBool(dados("Complementar"))
    
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

Function ConverteBRParaData(textoData As String) As Variant
    ' Converte string no formato DD/MM/YYYY para Date
    ' Garante que datas brasileiras sejam interpretadas corretamente
    
    Dim partes() As String
    Dim dia As Integer
    Dim mes As Integer
    Dim ano As Integer
    
    On Error Resume Next
    
    If Trim(textoData) = "" Then
        ConverteBRParaData = ""
        Exit Function
    End If
    
    ' Separa DD/MM/YYYY
    partes = Split(textoData, "/")
    
    If UBound(partes) = 2 Then
        dia = CInt(partes(0))
        mes = CInt(partes(1))
        ano = CInt(partes(2))
        
        ' Valida se é uma data válida
        If dia >= 1 And dia <= 31 And mes >= 1 And mes <= 12 And ano > 1900 Then
            ConverteBRParaData = DateSerial(ano, mes, dia)
            Exit Function
        End If
    End If
    
    ' Se falhou, tenta conversão direta (pode dar problema)
    ConverteBRParaData = textoData
    
    On Error GoTo 0
    
End Function

Sub CarregarVinculosCliente(ID_Cliente As Long)

    Dim ws As Worksheet
    Dim wsVinc As Worksheet
    Dim ultima As Long
    Dim i As Long
    Dim dataInicio As String
    Dim dataFim As String
    Dim tempoAnos As String
    Dim seq As String
    Dim observacoes As String

    Set ws = Sheets("Vinculos")
    
    ' Busca Seq no CSV via campo Observacoes
    Set wsVinc = ws

    frmCadastro.lstVinculos.Clear
    
    ' Configura ListBox - 6 colunas
    frmCadastro.lstVinculos.ColumnCount = 6
    ' Larguras: Seq(60) | Início(94) | Fim(94) | Tipo(360) | Tempo(70) | ID_Vinculo(0-oculto)
    frmCadastro.lstVinculos.ColumnWidths = "60 pt;94 pt;94 pt;360 pt;70 pt;0 pt"
    
    ' Adiciona linha de cabeçalho
    frmCadastro.lstVinculos.AddItem ""
    frmCadastro.lstVinculos.List(0, 0) = "Seq"
    frmCadastro.lstVinculos.List(0, 1) = "Início"
    frmCadastro.lstVinculos.List(0, 2) = "Fim"
    frmCadastro.lstVinculos.List(0, 3) = "Tipo de Filiado"
    frmCadastro.lstVinculos.List(0, 4) = "Tempo"
    
    ' Ajusta altura do ListBox dinamicamente
    On Error Resume Next
    frmCadastro.lstVinculos.IntegralHeight = False
    On Error GoTo 0

    ultima = ws.Cells(ws.Rows.count, 1).End(xlUp).Row

    For i = 2 To ultima
        If ws.Cells(i, 2).Value = ID_Cliente Then
            
            ' Formata datas explicitamente em formato brasileiro
            If IsDate(ws.Cells(i, 3).Value) Then
                dataInicio = Format(ws.Cells(i, 3).Value, "dd/mm/yyyy")
            Else
                dataInicio = ws.Cells(i, 3).Value
            End If
            
            If IsDate(ws.Cells(i, 4).Value) Then
                dataFim = Format(ws.Cells(i, 4).Value, "dd/mm/yyyy")
            Else
                dataFim = ws.Cells(i, 4).Value
            End If
            
            ' Calcula tempo em anos
            tempoAnos = ""
            If IsDate(ws.Cells(i, 3).Value) Then
                Dim dtIni As Date
                Dim dtFim As Date
                dtIni = ws.Cells(i, 3).Value
                
                If IsDate(ws.Cells(i, 4).Value) Then
                    dtFim = ws.Cells(i, 4).Value
                    
                    Dim anos As Double
                    anos = (dtFim - dtIni) / 365.25
                    
                    ' Só mostra se for maior que 0
                    If anos > 0.01 Then
                        tempoAnos = Format(anos, "0.00") & " anos"
                    End If
                End If
            End If
            
            ' Extrai Seq das observações ou usa contador sequencial
            seq = CStr(i - 1)
            
            ' Adiciona item: Seq | Início | Fim | Tipo | Tempo | ID_Vinculo(oculto)
            frmCadastro.lstVinculos.AddItem seq
            frmCadastro.lstVinculos.List(frmCadastro.lstVinculos.ListCount - 1, 1) = dataInicio
            frmCadastro.lstVinculos.List(frmCadastro.lstVinculos.ListCount - 1, 2) = dataFim
            frmCadastro.lstVinculos.List(frmCadastro.lstVinculos.ListCount - 1, 3) = ws.Cells(i, 5).Value
            frmCadastro.lstVinculos.List(frmCadastro.lstVinculos.ListCount - 1, 4) = tempoAnos
            frmCadastro.lstVinculos.List(frmCadastro.lstVinculos.ListCount - 1, 5) = ws.Cells(i, 1).Value ' ID_Vinculo oculto

        End If
    Next i
    
    ' Força o ListBox a mostrar o topo após carregar
    On Error Resume Next
    If frmCadastro.lstVinculos.ListCount > 0 Then
        frmCadastro.lstVinculos.TopIndex = 0
    End If
    On Error GoTo 0

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

    ' Primeiro detecta concomitância entre todos os vínculos
    Call DetectarConcomitancia(ID_Cliente)

    ' Varre todos os vínculos do cliente
    ultima = wsV.Cells(wsV.Rows.count, 1).End(xlUp).Row

    For i = 2 To ultima
        If wsV.Cells(i, 2).Value = ID_Cliente Then

            If wsV.Cells(i, 6).Value = "Sim" Then temEspecial = True
            If wsV.Cells(i, 11).Value = True Then temRural = True
            If wsV.Cells(i, 12).Value = True Then temMilitar = True
            If wsV.Cells(i, 13).Value = True Then temExterior = True
            If wsV.Cells(i, 14).Value = True Then temConcomitante = True
            If wsV.Cells(i, 15).Value = True Then temAtraso = True
            If wsV.Cells(i, 16).Value = True Then temComplementar = True

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

Sub DetectarConcomitancia(ID_Cliente As Long)
    ' Detecta automaticamente vínculos concomitantes (sobreposição de períodos)
    ' e marca o flag Concomitante em cada vínculo afetado
    
    Dim ws As Worksheet
    Dim ultima As Long
    Dim i As Long, j As Long
    Dim linha1 As Long, linha2 As Long
    Dim inicio1 As Date, fim1 As Date
    Dim inicio2 As Date, fim2 As Date
    
    Set ws = Sheets("Vinculos")
    
    ' Primeiro limpa o flag de todos os vínculos do cliente
    ultima = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    For i = 2 To ultima
        If ws.Cells(i, 2).Value = ID_Cliente Then
            ws.Cells(i, 14).Value = False  ' Coluna 14 = Concomitante
        End If
    Next i
    
    ' Compara todos os pares de vínculos
    For i = 2 To ultima
        If ws.Cells(i, 2).Value = ID_Cliente Then
            linha1 = i
            
            ' Valida datas do vínculo 1
            On Error Resume Next
            inicio1 = CDate(ws.Cells(linha1, 3).Value)
            fim1 = CDate(ws.Cells(linha1, 4).Value)
            On Error GoTo 0
            
            If inicio1 = 0 Or fim1 = 0 Then GoTo ProximoI
            
            ' Compara com todos os outros vínculos do mesmo cliente
            For j = i + 1 To ultima
                If ws.Cells(j, 2).Value = ID_Cliente Then
                    linha2 = j
                    
                    ' Valida datas do vínculo 2
                    On Error Resume Next
                    inicio2 = CDate(ws.Cells(linha2, 3).Value)
                    fim2 = CDate(ws.Cells(linha2, 4).Value)
                    On Error GoTo 0
                    
                    If inicio2 = 0 Or fim2 = 0 Then GoTo ProximoJ
                    
                    ' Verifica sobreposição: há concomitância se um período começa antes do outro terminar
                    If (inicio1 <= fim2) And (inicio2 <= fim1) Then
                        ' Marca ambos os vínculos como concomitantes (coluna 14)
                        ws.Cells(linha1, 14).Value = True
                        ws.Cells(linha2, 14).Value = True
                    End If
                    
ProximoJ:
                End If
            Next j
            
ProximoI:
        End If
    Next i
    
End Sub

' ============================================
' SalvarVinculoComExtras - Salva vínculo com campos extras (Seq, CodigoEmp)
' ============================================
Sub SalvarVinculoComExtras(dados As Collection, seq As String, codigoEmp As String)
    Dim ws As Worksheet
    Dim linha As Long
    Dim idV As Long
    Dim dataInicio As Variant
    Dim dataFim As Variant
    
    Set ws = Sheets("Vinculos")
    
    If dados("ID_Vinculo") = 0 Then
        linha = ProximaLinha(ws)
        idV = NovoID("Vinculos", 1)
    Else
        linha = BuscarLinhaVinculo(dados("ID_Vinculo"))
        idV = dados("ID_Vinculo")
    End If
    
    ' Converte datas do formato brasileiro (DD/MM/YYYY) corretamente
    dataInicio = ConverteBRParaData(dados("Inicio"))
    dataFim = ConverteBRParaData(dados("Fim"))
    
    ' Ordem correta: ID_Vinculo, ID_Cliente, Data_Inicio, Data_Fim, Tipo_Vinculo,
    '                Especial, Grau_Especial, Salario_Contribuicao, Observacoes, Data_Cadastro,
    '                Rural, Militar, Exterior, Concomitante, Atraso, Complementar, Código Emp.
    
    ws.Cells(linha, 1).Value = idV
    ws.Cells(linha, 2).Value = dados("ID_Cliente")
    
    ' Formatar datas como texto para evitar auto-conversão do Excel
    ws.Cells(linha, 3).NumberFormat = "@"
    ws.Cells(linha, 3).Value = CStr(dados("Inicio"))
    
    ws.Cells(linha, 4).NumberFormat = "@"
    ws.Cells(linha, 4).Value = CStr(dados("Fim"))
    
    ws.Cells(linha, 5).Value = dados("Tipo")
    ws.Cells(linha, 6).Value = dados("Especial")
    ws.Cells(linha, 7).Value = dados("Grau")
    ws.Cells(linha, 8).Value = dados("Salario")
    
    ' Garantir que Observações seja salvo como texto (coluna 9)
    ws.Cells(linha, 9).NumberFormat = "@"
    ws.Cells(linha, 9).Value = CStr(dados("Observacoes"))
    
    ws.Cells(linha, 10).Value = Date  ' Data_Cadastro
    
    ' Garantir booleanos explícitos (não converter para texto em português)
    ws.Cells(linha, 11).Value = CBool(dados("Rural"))
    ws.Cells(linha, 12).Value = CBool(dados("Militar"))
    ws.Cells(linha, 13).Value = CBool(dados("Exterior"))
    ws.Cells(linha, 14).Value = CBool(dados("Concomitante"))
    ws.Cells(linha, 15).Value = CBool(dados("Atraso"))
    ws.Cells(linha, 16).Value = CBool(dados("Complementar"))
    
    ' Dados extras do CNIS (colunas 17-18)
    ws.Cells(linha, 17).Value = codigoEmp    ' Código da Empresa (CNPJ)
    ws.Cells(linha, 18).Value = seq          ' Seq do CNIS
    
    Call AtualizarIndicadoresCliente(dados("ID_Cliente"))
End Sub

