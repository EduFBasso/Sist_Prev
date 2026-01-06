Attribute VB_Name = "ModClientes"
Option Explicit

' Módulo de Gerenciamento de Clientes
' Sistema de Planejamento Previdenciário RGPS

' Função para cadastrar novo cliente
Public Function CadastrarCliente(nome As String, cpf As String, dataNascimento As Date, _
                                 sexo As String, email As String, telefone As String) As Boolean
    
    On Error GoTo ErroHandler
    
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    
    ' Referência à planilha de clientes
    Set ws = ThisWorkbook.Worksheets("Clientes")
    
    ' Validar CPF
    If Not ValidarCPF(cpf) Then
        MsgBox "CPF inválido!", vbExclamation, "Erro de Validação"
        CadastrarCliente = False
        Exit Function
    End If
    
    ' Verificar se CPF já existe
    If CPFJaExiste(cpf) Then
        MsgBox "CPF já cadastrado no sistema!", vbExclamation, "Erro de Validação"
        CadastrarCliente = False
        Exit Function
    End If
    
    ' Encontrar próxima linha disponível
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Inserir dados do cliente
    ws.Cells(ultimaLinha, 1).Value = ultimaLinha - 1 ' ID
    ws.Cells(ultimaLinha, 2).Value = nome
    ws.Cells(ultimaLinha, 3).Value = cpf
    ws.Cells(ultimaLinha, 4).Value = dataNascimento
    ws.Cells(ultimaLinha, 5).Value = sexo
    ws.Cells(ultimaLinha, 6).Value = email
    ws.Cells(ultimaLinha, 7).Value = telefone
    ws.Cells(ultimaLinha, 8).Value = Now() ' Data de cadastro
    
    MsgBox "Cliente cadastrado com sucesso!", vbInformation, "Sucesso"
    CadastrarCliente = True
    Exit Function
    
ErroHandler:
    MsgBox "Erro ao cadastrar cliente: " & Err.Description, vbCritical, "Erro"
    CadastrarCliente = False
End Function

' Função para validar CPF
Public Function ValidarCPF(cpf As String) As Boolean
    Dim i As Integer
    Dim soma As Integer
    Dim resto As Integer
    Dim digito1 As Integer
    Dim digito2 As Integer
    Dim cpfNumeros As String
    
    ' Remove caracteres não numéricos
    cpfNumeros = ""
    For i = 1 To Len(cpf)
        If IsNumeric(Mid(cpf, i, 1)) Then
            cpfNumeros = cpfNumeros & Mid(cpf, i, 1)
        End If
    Next i
    
    ' Verificar se tem 11 dígitos
    If Len(cpfNumeros) <> 11 Then
        ValidarCPF = False
        Exit Function
    End If
    
    ' Verificar CPFs inválidos conhecidos
    If cpfNumeros = "00000000000" Or cpfNumeros = "11111111111" Or _
       cpfNumeros = "22222222222" Or cpfNumeros = "33333333333" Or _
       cpfNumeros = "44444444444" Or cpfNumeros = "55555555555" Or _
       cpfNumeros = "66666666666" Or cpfNumeros = "77777777777" Or _
       cpfNumeros = "88888888888" Or cpfNumeros = "99999999999" Then
        ValidarCPF = False
        Exit Function
    End If
    
    ' Calcular primeiro dígito verificador
    soma = 0
    For i = 1 To 9
        soma = soma + CInt(Mid(cpfNumeros, i, 1)) * (11 - i)
    Next i
    resto = (soma * 10) Mod 11
    If resto = 10 Then resto = 0
    digito1 = resto
    
    ' Calcular segundo dígito verificador
    soma = 0
    For i = 1 To 10
        soma = soma + CInt(Mid(cpfNumeros, i, 1)) * (12 - i)
    Next i
    resto = (soma * 10) Mod 11
    If resto = 10 Then resto = 0
    digito2 = resto
    
    ' Verificar dígitos
    If CInt(Mid(cpfNumeros, 10, 1)) = digito1 And CInt(Mid(cpfNumeros, 11, 1)) = digito2 Then
        ValidarCPF = True
    Else
        ValidarCPF = False
    End If
End Function

' Função para verificar se CPF já existe
Public Function CPFJaExiste(cpf As String) As Boolean
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim cpfExistente As String
    
    Set ws = ThisWorkbook.Worksheets("Clientes")
    ultimaLinha = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    
    For i = 2 To ultimaLinha
        cpfExistente = ws.Cells(i, 3).Value
        If cpfExistente = cpf Then
            CPFJaExiste = True
            Exit Function
        End If
    Next i
    
    CPFJaExiste = False
End Function

' Função para buscar cliente por CPF
Public Function BuscarClientePorCPF(cpf As String) As Long
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Worksheets("Clientes")
    ultimaLinha = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    
    For i = 2 To ultimaLinha
        If ws.Cells(i, 3).Value = cpf Then
            BuscarClientePorCPF = i
            Exit Function
        End If
    Next i
    
    BuscarClientePorCPF = 0
End Function

' Função para listar todos os clientes
Public Sub ListarClientes()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim mensagem As String
    
    Set ws = ThisWorkbook.Worksheets("Clientes")
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If ultimaLinha < 2 Then
        MsgBox "Nenhum cliente cadastrado.", vbInformation, "Lista de Clientes"
        Exit Sub
    End If
    
    mensagem = "CLIENTES CADASTRADOS:" & vbCrLf & vbCrLf
    
    For i = 2 To ultimaLinha
        mensagem = mensagem & "ID: " & ws.Cells(i, 1).Value & vbCrLf
        mensagem = mensagem & "Nome: " & ws.Cells(i, 2).Value & vbCrLf
        mensagem = mensagem & "CPF: " & ws.Cells(i, 3).Value & vbCrLf
        mensagem = mensagem & "Data Nasc.: " & Format(ws.Cells(i, 4).Value, "dd/mm/yyyy") & vbCrLf
        mensagem = mensagem & String(40, "-") & vbCrLf
    Next i
    
    MsgBox mensagem, vbInformation, "Lista de Clientes"
End Sub

' Inicializar planilha de clientes
Public Sub InicializarPlanilhaClientes()
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Clientes")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Clientes"
    End If
    
    ' Criar cabeçalhos se não existirem
    If ws.Cells(1, 1).Value = "" Then
        ws.Cells(1, 1).Value = "ID"
        ws.Cells(1, 2).Value = "Nome"
        ws.Cells(1, 3).Value = "CPF"
        ws.Cells(1, 4).Value = "Data Nascimento"
        ws.Cells(1, 5).Value = "Sexo"
        ws.Cells(1, 6).Value = "Email"
        ws.Cells(1, 7).Value = "Telefone"
        ws.Cells(1, 8).Value = "Data Cadastro"
        
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
