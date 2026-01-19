# CONFIGURAÇÃO DE DADOS DA EMPRESA

## 📋 Planilha Config_Empresa

Criar manualmente no Excel uma nova planilha com o nome **"Config_Empresa"** com os seguintes campos:

| Campo                    | Valor de Exemplo                            |
| ------------------------ | ------------------------------------------- |
| **Razao_Social**         | Advocacia Previdenciária Silva & Associados |
| **Nome_Fantasia**        | Silva Advocacia                             |
| **CNPJ**                 | 12.345.678/0001-90                          |
| **Endereco**             | Rua das Flores, 123 - Centro                |
| **Cidade**               | São Paulo                                   |
| **Estado**               | SP                                          |
| **CEP**                  | 01234-567                                   |
| **Telefone**             | (11) 3456-7890                              |
| **Email**                | contato@silvaadvocacia.com.br               |
| **Site**                 | www.silvaadvocacia.com.br                   |
| **Logo_Path**            | C:\Empresas\Logo.png                        |
| **OAB_Numero**           | OAB/SP 123.456                              |
| **Advogado_Responsavel** | Dr. João Silva                              |

## 🎨 Estrutura da planilha:

```
Coluna A: Nome do parâmetro
Coluna B: Valor do parâmetro
```

Exemplo:

```
A1: Parametro           B1: Valor
A2: Razao_Social        B2: Advocacia Previdenciária Silva & Associados
A3: Nome_Fantasia       B3: Silva Advocacia
A4: CNPJ                B4: 12.345.678/0001-90
...
```

## 📝 Código VBA para acessar:

```vb
Function GetDadosEmpresa(campo As String) As String
    Dim ws As Worksheet
    Dim i As Long

    On Error Resume Next
    Set ws = Sheets("Config_Empresa")
    On Error GoTo 0

    If ws Is Nothing Then
        GetDadosEmpresa = ""
        Exit Function
    End If

    ' Buscar campo na coluna A
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If UCase(Trim(ws.Cells(i, 1).Value)) = UCase(campo) Then
            GetDadosEmpresa = ws.Cells(i, 2).Value
            Exit Function
        End If
    Next i

    GetDadosEmpresa = ""
End Function
```

## 🖨️ Uso no Cabeçalho de Impressão:

```vb
Dim nomeEmpresa As String
Dim telefone As String
Dim email As String

nomeEmpresa = GetDadosEmpresa("Nome_Fantasia")
telefone = GetDadosEmpresa("Telefone")
email = GetDadosEmpresa("Email")

' Usar no cabeçalho do relatório
txtCabecalho.Value = nomeEmpresa & vbCrLf & telefone & " | " & email
```
