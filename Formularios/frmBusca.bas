' ============================================================
' Buscar registros na planilha Cadastro_Clientes
' ============================================================

Private Sub UserForm_Initialize()
    ' A validação principal já foi feita no frmPrincipal
    ' Aqui apenas carrega os dados
    Call cmdBuscar_Click
End Sub

Private Sub cmdBuscar_Click()

    Dim ws As Worksheet
    Dim ultima As Long
    Dim i As Long
    Dim termo As String
    Dim resultados() As Variant
    Dim count As Long

    termo = LCase(Trim(Me.txtFiltro.Value))
    
    On Error Resume Next
    Set ws = Sheets("Cadastro_Clientes")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Planilha 'Cadastro_Clientes' não encontrada.", vbCritical
        Exit Sub
    End If

    Me.lstResultados.Clear

    ultima = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    ' Verificar se há registros
    If ultima < 2 Then
        MsgBox "Nenhum cliente cadastrado.", vbInformation
        Exit Sub
    End If

    ' Criar matriz dinâmica
    ReDim resultados(1 To ultima - 1, 1 To 3)
    count = 0

    ' Coletar resultados
    For i = 2 To ultima
        If ws.Cells(i, 1).Value <> "" Then

            Dim nome As String
            Dim cpf As String

            nome = LCase(ws.Cells(i, 2).Value)
            cpf = LCase(ws.Cells(i, 3).Value)

            If InStr(nome, termo) > 0 Or InStr(cpf, termo) > 0 Then
                count = count + 1
                resultados(count, 1) = ws.Cells(i, 1).Value ' ID
                resultados(count, 2) = ws.Cells(i, 2).Value ' Nome
                resultados(count, 3) = ws.Cells(i, 3).Value ' CPF
            End If

        End If
    Next i

    ' Redimensionar para o tamanho real
    If count = 0 Then
        MsgBox "Nenhum resultado encontrado para: " & Me.txtFiltro.Value, vbInformation
        Exit Sub
    End If
    
    ReDim Preserve resultados(1 To count, 1 To 3)

    ' Ordenar por nome (coluna 2)
    Call OrdenarResultados(resultados, 2)

    ' Preencher ListBox
    For i = 1 To count
        Me.lstResultados.AddItem resultados(i, 1)
        Me.lstResultados.List(Me.lstResultados.ListCount - 1, 1) = resultados(i, 2)
        Me.lstResultados.List(Me.lstResultados.ListCount - 1, 2) = resultados(i, 3)
    Next i

End Sub

' ============================================================
' Selecionar registro e carregar no frmCadastro
' ============================================================
Private Sub cmdSelecionar_Click()

    If Me.lstResultados.ListIndex = -1 Then
        MsgBox "Selecione um registro.", vbExclamation
        Exit Sub
    End If

    Dim ID As Long
    ID = Me.lstResultados.List(Me.lstResultados.ListIndex, 0)

    Call CarregarCliente(ID)

    frmCadastro.Show
    Unload Me

End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub

Private Sub lstResultados_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.lstResultados.ListIndex <> -1 Then
        Call cmdSelecionar_Click
    End If
End Sub
