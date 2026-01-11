# Passo a Passo - Simulação Completa do Cliente João Carlos

## 📋 DADOS DO CLIENTE EXTRAÍDO

**Nome:** JOAO CARLOS EDUARDO FIGUEIREDO BASSO  
**CPF:** 170.140.798-16  
**NIT:** 125.37781.66-1  
**Data Nascimento:** 22/03/1973  
**Idade Atual (11/01/2026):** 52 anos e 10 meses

## 🔄 PASSO 1: ABRIR O EXCEL E PREPARAR

1. Abra o arquivo **erp_prev.xlsm**
2. **Habilite as macros** quando solicitado
3. Verifique se as seguintes planilhas existem:
   - `Cadastro_Clientes`
   - `Vinculos`
   - `Config_Regras`
   - `Simulacoes`

## 👤 PASSO 2: CADASTRAR O CLIENTE MANUALMENTE

Como ainda não temos o formulário de importação automática do cabeçalho, vamos cadastrar manualmente:

### No Excel (planilha Cadastro_Clientes):

| Coluna | Campo       | Valor                                |
| ------ | ----------- | ------------------------------------ |
| A      | ID          | 1                                    |
| B      | Nome        | JOAO CARLOS EDUARDO FIGUEIREDO BASSO |
| C      | CPF         | 170.140.798-16                       |
| D      | PIS         | 125.37781.66-1                       |
| E      | RG          | (deixar vazio)                       |
| F      | Orgao       | (deixar vazio)                       |
| G      | Nascimento  | 22/03/1973                           |
| H      | Sexo        | M                                    |
| I      | EstadoCivil | (deixar vazio)                       |
| J      | Telefone    | (deixar vazio)                       |
| K      | Celular     | (deixar vazio)                       |
| L      | Email       | (deixar vazio)                       |

**OU via VBA:**

Pressione `Alt+F11` e execute no **Immediate Window** (Ctrl+G):

```vb
Sub CadastrarJoaoCarlos()
    Dim dados As New Collection

    dados.Add 1, "ID"
    dados.Add "JOAO CARLOS EDUARDO FIGUEIREDO BASSO", "Nome"
    dados.Add "170.140.798-16", "CPF"
    dados.Add "125.37781.66-1", "PIS"
    dados.Add "", "RG"
    dados.Add "", "Orgao"
    dados.Add #3/22/1973#, "Nascimento"
    dados.Add "M", "Sexo"
    dados.Add "", "EstadoCivil"
    dados.Add "", "Telefone"
    dados.Add "", "Celular"
    dados.Add "", "Email"
    dados.Add "", "CEP"
    dados.Add "", "Endereco"
    dados.Add "", "Numero"
    dados.Add "", "Complemento"
    dados.Add "", "Bairro"
    dados.Add "", "Cidade"
    dados.Add "", "UF"
    dados.Add "", "Filiacao"
    dados.Add "Empregado", "TipoSegurado"
    dados.Add False, "Especial"
    dados.Add False, "Rural"
    dados.Add False, "Militar"
    dados.Add False, "Exterior"
    dados.Add False, "Concomitante"
    dados.Add False, "Atraso"
    dados.Add False, "Complementar"
    dados.Add "", "Observacoes"

    Call SalvarCliente(dados)
    MsgBox "Cliente cadastrado com sucesso!"
End Sub
```

## 📊 PASSO 3: IMPORTAR VÍNCULOS

### Via VBA (Recomendado):

No **Immediate Window** (Alt+F11 → Ctrl+G):

```vb
Call ImportarVinculosDeCSV("C:\CAMINHO_COMPLETO\saida_cnis_vinculos_estruturado.csv", 1)
```

**IMPORTANTE:** Substitua `C:\CAMINHO_COMPLETO\` pelo caminho real onde está o arquivo CSV.

### Vínculos que Serão Importados:

Com base no CSV, o João Carlos tem **10 vínculos**:

1. EMBIARA SERVICOS (19/01/1995 - 02/06/1995) = 4,5 meses
2. DANTEK (02/10/1995 - até 11/1998) = ~3 anos
3. DANTEK (01/05/1996 - 14/10/1996) = 5,5 meses
4. DANTEK (10/05/1996 - 10/1996) = 5 meses
5. DANTEK (01/07/1996 - 14/10/1996) = 3,5 meses
6. PTT EMPREENDIMENTOS (06/01/1998 - 05/04/1998) = 3 meses
7. ZABECCA (03/11/1998 - 29/05/1999) = 7 meses
8. DIMARES (01/06/1999 - 02/07/2002) = 3 anos
9. JR EQUIPAMENTOS (03/07/2002 - 15/03/2004) = 1 ano e 8 meses
10. JR EQUIPAMENTOS (01/07/2004 - 02/05/2008) = 3 anos e 10 meses

**Observação:** Há sobreposição de vínculos DANTEK (2, 3, 4, 5) - o sistema detectará automaticamente!

## 🔍 PASSO 4: VERIFICAR A IMPORTAÇÃO

Execute no Immediate Window:

```vb
Sub VerificarImportacao()
    Dim ws As Worksheet
    Dim ultima As Long
    Dim i As Long
    Dim count As Long

    Set ws = Sheets("Vinculos")
    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    count = 0
    For i = 2 To ultima
        If ws.Cells(i, 2).Value = 1 Then
            count = count + 1
            Debug.Print "Vínculo " & count & ": " & ws.Cells(i, 3).Value & " até " & ws.Cells(i, 4).Value
        End If
    Next i

    MsgBox "Total de vínculos importados: " & count
End Sub
```

## ⏱️ PASSO 5: CALCULAR TEMPO TOTAL

Execute no Immediate Window:

```vb
Sub CalcularTempoJoao()
    Dim tempo As Double
    Dim tempoEspecial As Double
    Dim idade As Long
    Dim pontos As Double

    tempo = CalcularTempo(1)
    tempoEspecial = CalcularTempoEspecial(1)
    idade = CalcularIdade(#3/22/1973#)
    pontos = idade + tempo

    Debug.Print "Idade atual: " & idade & " anos"
    Debug.Print "Tempo total: " & Format(tempo, "0.00") & " anos"
    Debug.Print "Tempo especial: " & Format(tempoEspecial, "0.00") & " anos"
    Debug.Print "Pontos atuais: " & Format(pontos, "0.00")

    MsgBox "Tempo total: " & Format(tempo, "0.00") & " anos" & vbCrLf & _
           "Idade: " & idade & " anos" & vbCrLf & _
           "Pontos: " & Format(pontos, "0.00")
End Sub
```

### Resultado Esperado:

- **Idade:** 52 anos
- **Tempo Total:** ~13-15 anos (considerando sobreposições)
- **Pontos:** ~65-67

## 🎯 PASSO 6: SIMULAR TODAS AS REGRAS

Execute no Immediate Window:

```vb
Sub SimularTodasRegras()
    Dim r As Collection
    Dim regra As String
    Dim msg As String

    ' Regra 1: Tempo de Contribuição
    Set r = RegraTempoContribuicao(1)
    msg = "=== TEMPO DE CONTRIBUIÇÃO ===" & vbCrLf
    msg = msg & "Direito: " & r("Direito") & vbCrLf
    msg = msg & "Falta: " & Format(r("Falta"), "0.00") & " anos" & vbCrLf
    msg = msg & "Data Provável: " & r("DataPrevista") & vbCrLf & vbCrLf

    ' Regra 2: Idade
    Set r = RegraIdade(1)
    msg = msg & "=== APOSENTADORIA POR IDADE ===" & vbCrLf
    msg = msg & "Direito: " & r("Direito") & vbCrLf
    msg = msg & "Falta: " & Format(r("Falta"), "0.00") & " anos" & vbCrLf
    msg = msg & "Data Provável: " & r("DataPrevista") & vbCrLf & vbCrLf

    ' Regra 3: Pontos
    Set r = RegraPontos(1)
    msg = msg & "=== REGRA DE PONTOS ===" & vbCrLf
    msg = msg & "Direito: " & r("Direito") & vbCrLf
    msg = msg & "Pontos atuais: " & Format(r("Pontos"), "0.00") & vbCrLf
    msg = msg & "Falta: " & Format(r("Falta"), "0.00") & " anos" & vbCrLf
    msg = msg & "Data Provável: " & r("DataPrevista") & vbCrLf & vbCrLf

    ' Regra 4: Pedágio 50%
    Set r = RegraPedagio50(1)
    msg = msg & "=== PEDÁGIO 50% ===" & vbCrLf
    msg = msg & "Direito: " & r("Direito") & vbCrLf
    msg = msg & "Falta: " & Format(r("Falta"), "0.00") & " anos" & vbCrLf
    msg = msg & "Obs: " & r("Obs") & vbCrLf & vbCrLf

    ' Regra 5: Pedágio 100%
    Set r = RegraPedagio100(1)
    msg = msg & "=== PEDÁGIO 100% ===" & vbCrLf
    msg = msg & "Direito: " & r("Direito") & vbCrLf
    msg = msg & "Falta: " & Format(r("Falta"), "0.00") & " anos" & vbCrLf
    msg = msg & "Obs: " & r("Obs") & vbCrLf

    Debug.Print msg
    MsgBox msg, vbInformation, "Simulação - João Carlos"
End Sub
```

## 🏆 PASSO 7: ANÁLISE AUTOMÁTICA DA MELHOR REGRA

Execute no Immediate Window:

```vb
Sub AnalisarMelhorOpcao()
    Dim resultado As Collection
    Dim melhorRegra As String
    Dim msg As String

    Set resultado = AnalisarMelhorRegra(1)
    melhorRegra = resultado("MelhorRegra")

    msg = "🎯 MELHOR OPÇÃO PARA JOÃO CARLOS" & vbCrLf & vbCrLf

    Select Case melhorRegra
        Case "TEMPO"
            msg = msg & "Regra: Aposentadoria por Tempo de Contribuição"
        Case "IDADE"
            msg = msg & "Regra: Aposentadoria por Idade"
        Case "PONTOS"
            msg = msg & "Regra: Pontos"
        Case "PEDAGIO50"
            msg = msg & "Regra: Pedágio 50%"
        Case "PEDAGIO100"
            msg = msg & "Regra: Pedágio 100%"
    End Select

    msg = msg & vbCrLf & vbCrLf
    msg = msg & "Direito: " & resultado("Direito") & vbCrLf
    msg = msg & "Falta: " & Format(resultado("Falta"), "0.00") & " anos" & vbCrLf
    msg = msg & "Data Provável: " & resultado("DataPrevista") & vbCrLf
    msg = msg & "Idade Projetada: " & resultado("IdadeProjetada") & " anos" & vbCrLf
    msg = msg & "Observação: " & resultado("Obs")

    Debug.Print msg
    MsgBox msg, vbInformation, "Melhor Opção"
End Sub
```

## 💰 PASSO 8: CALCULAR VALOR ESTIMADO DO BENEFÍCIO

Execute no Immediate Window:

```vb
Sub CalcularValorBeneficioJoao()
    Dim tempo As Double
    Dim valor As Double
    Dim msg As String

    tempo = CalcularTempo(1)

    ' Valor atual
    valor = CalcularValorBeneficio(1, tempo)
    msg = "💰 VALOR ESTIMADO DO BENEFÍCIO" & vbCrLf & vbCrLf
    msg = msg & "Tempo atual: " & Format(tempo, "0.00") & " anos" & vbCrLf
    msg = msg & "Valor estimado hoje: R$ " & Format(valor, "#,##0.00") & vbCrLf & vbCrLf

    ' Valor projetado (quando completar 35 anos)
    valor = CalcularValorBeneficio(1, 35)
    msg = msg & "Se completar 35 anos:" & vbCrLf
    msg = msg & "Valor estimado: R$ " & Format(valor, "#,##0.00")

    Debug.Print msg
    MsgBox msg, vbInformation, "Valor do Benefício"
End Sub
```

## 📊 RESULTADO ESPERADO

Com base nos vínculos extraídos (1995-2008), João Carlos deve ter:

### Situação Atual:

- ❌ **Não tem direito** a nenhuma regra ainda
- 📊 **Tempo contribuído:** ~13-15 anos
- 🎂 **Idade:** 52 anos
- 🎯 **Pontos:** ~65-67

### O que falta:

- **Para Tempo (35 anos):** ~20-22 anos
- **Para Idade (65 anos):** ~13 anos
- **Para Pontos (105):** ~38-40 pontos (≈19-20 anos)

### Melhor opção:

Provavelmente **IDADE** será a melhor opção (falta menos tempo)

## ⚠️ OBSERVAÇÕES IMPORTANTES

### 1. Vínculos Sobrepostos (DANTEK)

O João Carlos tem 4 vínculos da DANTEK que se sobrepõem:

- 02/10/1995 - 11/1998
- 01/05/1996 - 14/10/1996
- 10/05/1996 - 10/1996
- 01/07/1996 - 14/10/1996

O sistema **detectará automaticamente** e contará apenas o período único.

### 2. Vínculos Não Capturados

Como você mencionou, há outros tipos de vínculos que não foram extraídos:

- **Seq. 11:** Recolhimento Facultativo (09/2019 - 10/2019)
- **Seq. 12:** Contribuições Facultativas (2024-2025)

Vamos precisar **atualizar o script** para capturar esses também!

### 3. Valores Sem Correção

Os valores calculados são estimativas e **não incluem correção monetária**.

## 🔄 PRÓXIMO PASSO: CAPTURAR VÍNCULOS FACULTATIVOS

Você quer que eu atualize o script Python para capturar também:

1. Vínculos de tipo "Facultativo" e "Recolhimento"
2. Tabelas de "Contribuições" (sem empresa)
3. Campos: Competência, Data Pgto, Contribuição, Salário Contribuição, Indicadores

Isso adicionará mais tempo de contribuição ao João Carlos (2019 e 2024-2025)!

---

**Execute os passos acima e me informe os resultados!**  
Depois ajustaremos o script para capturar os outros vínculos. 🚀
