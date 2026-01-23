# MAPEAMENTO DO SISTEMA VBA - ERP PREVIDENCIÁRIO

## 📁 ESTRUTURA DO SISTEMA

### **1. Formulários (UserForms)**
```
Formularios/
├── frmBusca.bas          - Interface de busca de clientes
├── frmCadastro.bas       - Cadastro de clientes e dados pessoais
├── frmPrincipal.bas      - Tela principal do sistema
├── frmRelatorio.bas      - Geração de relatórios
├── frmSimulacoes.bas     - 🎯 SIMULAÇÕES DE APOSENTADORIA
└── frmVinculos.bas       - Cadastro de vínculos e períodos
```

### **2. Módulos (Lógica de Negócio)**
```
Modulos/
├── modBusca.bas          - Funções de pesquisa
├── modCadastro.bas       - Lógica de cadastro
├── modDB.bas             - Acesso ao banco (planilhas)
├── modDocumentos.bas     - Geração de documentos
├── modImportacao.bas     - Importação do CNIS
├── modSimulacoes.bas     - 🎯 MOTOR DE CÁLCULO
├── modUtil.bas           - Utilitários gerais
└── modVinculos.bas       - Processamento de vínculos
```

### **3. Planilhas (Base de Dados + Regras)**
```
Planilhas/
└── Config_Regras.bas     - 🎯 REGRAS PREVIDENCIÁRIAS
    ├── Regras de Aposentadoria
    ├── Fatores de Conversão
    ├── Coeficientes
    └── Configurações do Sistema

Outras planilhas Excel:
├── Clientes              - Cadastro de clientes
├── Vinculos              - Períodos de contribuição
├── Remuneracoes          - Salários mensais
└── Simulacoes            - Resultados de simulações
```

---

## 🎯 FLUXO DE CÁLCULO DE APOSENTADORIA

### **Sequência de Execução:**

```
1. frmSimulacoes (Interface)
   ↓
2. modSimulacoes.CalcularAposentadoria()
   ↓
3. modVinculos.ProcessarVinculos()
   ├─ Ler vinculos da planilha
   ├─ Verificar concomitância
   ├─ Converter atividade especial
   └─ Somar tempo total
   ↓
4. Config_Regras (Buscar regras aplicáveis)
   ├─ Idade mínima
   ├─ Tempo mínimo
   ├─ Fator de conversão especial
   └─ Coeficiente
   ↓
5. modSimulacoes.CalcularCoeficiente()
   ↓
6. modSimulacoes.CalcularMediaSalarial()
   ↓
7. Retornar resultado para frmSimulacoes
```

---

## 🔍 PONTOS CRÍTICOS IDENTIFICADOS

### **❌ Problema 1: CONCOMITÂNCIA**
**Local:** `modVinculos.bas` - função `SomarTempoContribuicao()`

**Esperado:**
```vba
Function SomarTempoContribuicao() As Long
    ' Deve eliminar períodos sobrepostos
    ' Vínculo 2: 02/10/1995 - 30/11/1998
    ' Vínculo 3: 01/05/1996 - 14/10/1996 (dentro do período do vínculo 2)
    ' Resultado: Não somar vínculo 3 (já está coberto pelo 2)
End Function
```

**Verificar:**
- Se há função `EliminarConcomitancia()`
- Se a flag `Concomitante=VERDADEIRO` está sendo usada
- Como os períodos sobrepostos são tratados

---

### **❌ Problema 2: ATIVIDADE ESPECIAL**
**Local:** `modVinculos.bas` - função `ConverterAtividadeEspecial()`

**Esperado:**
```vba
Function ConverterAtividadeEspecial(dias As Long, grau As Integer) As Long
    Select Case grau
        Case 15: ConverterAtividadeEspecial = dias * 2.33  ' 133% mais
        Case 20: ConverterAtividadeEspecial = dias * 1.75  ' 75% mais
        Case 25: ConverterAtividadeEspecial = dias * 1.40  ' 40% mais
        Case Else: ConverterAtividadeEspecial = dias
    End Select
End Function
```

**Verificar:**
- Se há conversão quando `Especial=Sim`
- Se usa o campo `Grau_Especial` (15, 20 ou 25)
- Como os dias convertidos são somados ao tempo comum

---

### **❌ Problema 3: IDADE MÍNIMA**
**Local:** `modSimulacoes.bas` - função `VerificarRequisitos()`

**Esperado:**
```vba
Function PodeAposentar(idadeAnos As Integer, tempoAnos As Double) As Boolean
    ' Reforma 2019 (Lei 13.982)
    Dim idadeMinima As Integer
    Dim tempoMinimo As Integer
    
    ' Buscar da planilha Config_Regras
    idadeMinima = 65  ' Homem (ou 62 para mulher)
    tempoMinimo = 15
    
    ' AMBOS os requisitos são obrigatórios
    If idadeAnos >= idadeMinima And tempoAnos >= tempoMinimo Then
        PodeAposentar = True
    Else
        PodeAposentar = False
    End If
End Function
```

**Verificar:**
- Se valida idade E tempo simultaneamente
- Se está usando apenas tempo (erro!)
- Qual regra está aplicando (Reforma 2019 ou anterior)

---

## 📊 PLANILHA CONFIG_REGRAS (Estrutura Esperada)

```
Aba: Regras_Aposentadoria
┌────────────────┬──────────┬──────────────┬─────────────┐
│ Tipo_Regra     │ Sexo     │ Idade_Minima │ Tempo_Min   │
├────────────────┼──────────┼──────────────┼─────────────┤
│ Por Idade      │ M        │ 65           │ 15          │
│ Por Idade      │ F        │ 62           │ 15          │
│ Por Pontos     │ M        │ -            │ 35          │
│ Por Pontos     │ F        │ -            │ 30          │
└────────────────┴──────────┴──────────────┴─────────────┘

Aba: Fatores_Conversao
┌─────────────┬───────────────┐
│ Grau_Esp    │ Fator         │
├─────────────┼───────────────┤
│ 15          │ 2.33          │
│ 20          │ 1.75          │
│ 25          │ 1.40          │
└─────────────┴───────────────┘

Aba: Coeficientes
┌─────────────┬──────────────┐
│ Anos_Contrib│ Coeficiente  │
├─────────────┼──────────────┤
│ 15          │ 60%          │
│ 20          │ 60%          │
│ 25          │ 70%          │
│ 30          │ 80%          │
│ 35          │ 90%          │
│ 40          │ 100%         │
└─────────────┴──────────────┘

Fórmula: 60% + 2% × MAX(0, Anos - 20)
```

---

## 🎯 PRÓXIMOS PASSOS

### **1. Análise do Código VBA**
- [ ] Ler `modSimulacoes.bas` linha por linha
- [ ] Identificar função principal de cálculo
- [ ] Mapear todas as variáveis e constantes
- [ ] Documentar lógica de negócio

### **2. Validação com Planilha Manual**
- [ ] Criar planilha Excel com cálculo correto
- [ ] Preencher com dados do João Carlos
- [ ] Calcular tempo SEM concomitância
- [ ] Converter atividade especial
- [ ] Comparar resultado VBA × Planilha

### **3. Correção do VBA**
- [ ] Implementar eliminação de concomitância
- [ ] Adicionar conversão de atividade especial
- [ ] Corrigir validação de idade mínima
- [ ] Testar com casos reais
- [ ] Validar com advogados

---

## 📝 OBSERVAÇÕES

**Data deste mapeamento:** 22/01/2026

**Caso de teste:** João Carlos Eduardo Figueiredo Basso
- 13 vínculos (1 especial, 5 concomitantes)
- Idade: 52 anos
- Sistema atual: "Pode aos 53 anos" ❌ INCORRETO
- Correto: "Pode aos 65 anos com tempo atual"

**Impacto:** Sistema está dando esperança falsa aos clientes!

---

## 🔗 ARQUIVOS RELACIONADOS

- `MAPEAMENTO_CODIGO.md` - Mapa geral do código
- `MAPEAMENTO_FRMSIMULACOES.md` - Detalhes do formulário
- `MOTOR_CALCULO_STATUS.md` - Status do motor de cálculo
- `VINCULOS_FACULTATIVOS_PENDENTE.md` - Questões pendentes
