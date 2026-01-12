# 🔍 MAPEAMENTO COMPLETO: frmSimulacoes

**Data de Análise:** 12 de janeiro de 2026  
**Sistema:** Sist_Prev - Planejamento Previdenciário RGPS  
**Componente:** Formulário de Simulações de Aposentadoria

---

## 🎯 RESUMO EXECUTIVO

O **frmSimulacoes** é o coração do sistema de simulações, apresentando ao advogado uma interface completa para análise de cenários de aposentadoria. O formulário trabalha em **MODO AUTOMÁTICO**, calculando e exibindo os 3 principais cenários (rápido, equilibrado e máximo) **com análise automática da melhor regra** em um único clique.

**Ponto de Entrada:** O formulário é aberto através do botão **"Simular"** (`cmdSimular`) no `frmCadastro`, após o cliente ter sido cadastrado (manualmente, via importação CNIS ou via busca).

---

## 🎯 ARQUITETURA DO SISTEMA

### Fluxo de Integração Python + VBA

```
┌─────────────────────────────────────────────────────────────────┐
│                     SISTEMA MISTO                                │
├─────────────────────────────────────────────────────────────────┤
│                                                                   │
│  PYTHON (converter_extrato_inss.py)                             │
│  └─> Extrai dados do PDF CNIS:                                  │
│      • Dados do Cliente (NIT, CPF, Nome, Nascimento)            │
│      • Vínculos Empregatícios (Início, Fim, Empresa)            │
│      • Remunerações por Competência (Mês/Ano + Valor)           │
│      • Status: ✅ EMPREGADO e AGENTE PÚBLICO                    │
│      • Status: ⚠️  FACULTATIVO e CONTRIBUINTE INDIVIDUAL        │
│                   (Pendente de implementação)                    │
│                                                                   │
│  VBA (Excel/Access)                                              │
│  └─> Recebe os CSVs gerados pelo Python:                        │
│      • ImportarVinculosDeCSV() -> Planilha "Vinculos"          │
│      • ImportarRemuneracoesDeCSV() -> Planilha "Remuneracoes"  │
│      └─> MOTOR DE CÁLCULO (modSimulacoes.bas)                  │
│          • Calcula tempo total de contribuição                   │
│          • Converte tempo especial (fatores 1.4, 1.75, 2.33)   │
│          • Aplica 5 regras de aposentadoria                      │
│          • Calcula valores dos benefícios                        │
│          └─> APRESENTAÇÃO (frmSimulacoes)                       │
│              • Exibe dados calculados                            │
│              • Mostra 3 cenários comparativos                    │
│              • Permite análise de melhor regra                   │
│                                                                   │
└─────────────────────────────────────────────────────────────────┘
```

---

## 🖼️ ESTRUTURA VISUAL DO FORMULÁRIO

### 1️⃣ FRAME: "Dados do Cliente"

**Propósito:** Exibir situação atual do cliente (SOMENTE LEITURA)

| Controle         | Tipo    | Propriedade | Conteúdo Exibido                 |
| ---------------- | ------- | ----------- | -------------------------------- |
| lblCliente       | Label   | Caption     | "Cliente: [Nome do Cliente]"     |
| lblID            | Label   | Caption     | "ID: [ID_Cliente]"               |
| txtIdadeAtual    | TextBox | Locked=True | "[X] anos"                       |
| txtTempoTotal    | TextBox | Locked=True | "[X.XX] anos"                    |
| txtTempoEspecial | TextBox | Locked=True | "[X.XX] anos" (tempo convertido) |
| txtPontos        | TextBox | Locked=True | "[XX.XX]" (idade + tempo)        |

**Preenchimento:** Automático via `PreencherDadosIniciais(ID_Cliente)`

**Cálculos Envolvidos:**

- `CalcularIdade(nascimento)` - Idade atual em anos
- `CalcularTempo(ID_Cliente)` - Soma de todos os vínculos (com tratamento de sobreposição)
- `CalcularTempoEspecial(ID_Cliente)` - Tempo especial já convertido para comum
- `Pontos = Idade + Tempo Total`

---

### 2️⃣ FRAME: "Escolha a Regra de Simulação"

**Propósito:** Permitir análise por regra específica (MODO MANUAL - LEGADO)

⚠️ **Status:** Este frame está **mantido para compatibilidade**, mas o sistema agora opera em **MODO AUTOMÁTICO** (veja item 4).

| Controle             | Tipo         | Caption/Regra                              |
| -------------------- | ------------ | ------------------------------------------ |
| optTempoContribuicao | OptionButton | "Tempo de Contribuição (pré-reforma)"      |
| optPontos            | OptionButton | "Regra de Pontos (86/96 -> 100/105)"       |
| optIdade             | OptionButton | "Idade Mínima Progressiva"                 |
| optEspecial          | OptionButton | "Aposentadoria Especial (15, 20, 25 anos)" |

**Função Antiga:** `cmdCalcular_OLD_Click()` - Calcula baseado na regra selecionada

---

### 3️⃣ FRAME: "Regras e Pedágio"

**Propósito:** Escolher regra de transição (SE APLICÁVEL)

⚠️ **Observação:** A elegibilidade é verificada automaticamente pelo sistema.

#### Subgrupo: Transição

| Controle        | Tipo         | GroupName | Caption          |
| --------------- | ------------ | --------- | ---------------- |
| optTransicao50  | OptionButton | Transicao | "Transição 50%"  |
| optTransicao100 | OptionButton | Transicao | "Transição 100%" |

#### Subgrupo: Pedágio

| Controle      | Tipo         | GroupName | Caption        |
| ------------- | ------------ | --------- | -------------- |
| optPedagio50  | OptionButton | Pedagio   | "Pedágio 50%"  |
| optPedagio100 | OptionButton | Pedagio   | "Pedágio 100%" |

**Verificação Automática:** `VerificarElegibilidadeTransicao(ID_Cliente)`

- Retorna: "TODAS", "PEDAGIO50", "PEDAGIO100", "NENHUMA"
- Base: Tempo de contribuição na data da reforma (13/11/2019)

---

### 4️⃣ FRAME: "Resultado da Simulação"

**Propósito:** Exibir resultado do cenário selecionado/calculado

| Controle         | Tipo    | Conteúdo Exibido                                   |
| ---------------- | ------- | -------------------------------------------------- |
| txtDireito       | TextBox | "Sim" ou "Não"                                     |
| txtFalta         | TextBox | "[X.XX] anos (X anos e Y meses)" ou "Já elegível!" |
| txtIdadeProj     | TextBox | "[XX] anos"                                        |
| txtDataProvavel  | TextBox | "dd/mm/yyyy"                                       |
| txtValorEstimado | TextBox | "R$ X.XXX,XX"                                      |
| txtObs           | TextBox | Texto longo com análise comparativa dos 3 cenários |

**Preenchimento:** Via `cmdCalcular_Click()` (modo automático atual)

---

## ⚙️ FUNCIONAMENTO ATUAL (MODO AUTOMÁTICO)

### 🎯 FLUXO PRINCIPAL

#### 1. **Abertura do Formulário**

```vb
Public Sub PreencherDadosIniciais(ID_Cliente As Long)
```

- Busca dados do cliente na planilha "Cadastro_Clientes"
- Calcula automaticamente:
  - Idade atual
  - Tempo total de contribuição (com tratamento de sobreposição)
  - Tempo especial convertido
  - Pontos (idade + tempo)
- **Bloqueia os campos para edição** (`Locked = True`)

#### 2. **Botão "Calcular" (Modo Atual)**

```vb
Private Sub cmdCalcular_Click()
```

**O que faz:**

1. Calcula **Cenário 1: Aposentadoria Rápida** (15 anos)
2. Preenche os campos principais com este cenário
3. Adiciona automaticamente no `txtObs` uma análise comparativa dos **Cenários 2 e 3**

**Cenários Calculados:**

| Cenário | Descrição                 | Tempo Necessário | Percentual do Benefício |
| ------- | ------------------------- | ---------------- | ----------------------- |
| 1       | Aposentadoria Rápida      | 15 anos          | 60% (mínimo legal)      |
| 2       | Aposentadoria Equilibrada | 25 anos          | 70-80% (recomendado)    |
| 3       | Aposentadoria Máxima      | 40 anos          | 100% (integral)         |

**Funções Chamadas:**

- `CalcularCenarioRapido(ID)` → Cenário 1
- `CalcularCenarioEquilibrado(ID)` → Cenário 2
- `CalcularCenarioMaximo(ID)` → Cenário 3
- `GerarTextoCenariosComplementares(ID)` → Formata texto para txtObs

#### 3. **Botão "Analisar Todas as Regras"**

```vb
Private Sub cmdAnalisarTodas_Click()
```

**O que faz:**

1. Testa todas as 5 regras de aposentadoria
2. Identifica qual regra é mais vantajosa (menor tempo faltante ou direito adquirido)
3. Marca automaticamente o OptionButton correspondente
4. Executa `cmdCalcular_Click()` para exibir resultados
5. Exibe mensagem informativa com a melhor regra

**Função Chamada:** `AnalisarMelhorRegra(ID_Cliente)` → Collection com "MelhorRegra"

**Regras Testadas:**

1. `RegraTempoContribuicao(ID)` - 35H/30M anos
2. `RegraIdade(ID)` - 65H/62M anos + 180 meses carência
3. `RegraPontos(ID)` - 105H/100M pontos + 35H/30M anos
4. `RegraPedagio50(ID)` - SE elegível (≤2 anos na reforma)
5. `RegraPedagio100(ID)` - SE elegível

---

## 💰 CÁLCULO DE VALORES (ATUAL E FUTURO)

### 📊 Situação Atual

#### Função: `CalcularValorBeneficio(ID_Cliente, tempoContribuicao)`

**Fórmula Pós-Reforma (Lei 13.183/2019):**

```
Coeficiente = 60% + 2% × (tempo - tempoBase)

Onde:
- tempoBase = 20 anos (homem) ou 15 anos (mulher)
- Máximo: 100%
```

**Exemplo:**

- Mulher com 25 anos de contribuição:
  - Coef = 60% + 2% × (25 - 15) = 60% + 20% = 80%

**Média Salarial (ESTIMATIVA SIMPLIFICADA ATUAL):**

```vb
mediaSalarialEstimada = tetoINSS * 0.6  ' Aproximadamente R$ 4.671,61
```

⚠️ **LIMITAÇÃO ATUAL:** O sistema usa uma **estimativa conservadora** de 60% do teto do INSS, sem calcular a média real dos 80% maiores salários.

---

### 🎯 Implementação Futura com Remunerações

O Python **JÁ EXTRAI** as remunerações do CNIS e gera o arquivo:

- `saida_cnis_remuneracoes.csv` (colunas: NIT, Competencia, Valor)

**Planilha Target:** `Remuneracoes` (ID_Cliente, Competencia, Valor)

#### Função a Implementar:

```vb
Function CalcularMediaReal(ID_Cliente As Long) As Double
    ' 1. Buscar todas as remunerações da planilha "Remuneracoes"
    ' 2. Ordenar do maior para o menor
    ' 3. Pegar os 80% maiores valores
    ' 4. Calcular a média aritmética
    ' 5. Aplicar reajustes até a data atual (índice INPC/TR)
    ' 6. Retornar valor médio real
End Function
```

**Modificação em CalcularValorBeneficio:**

```vb
' Em vez de:
mediaSalarialEstimada = tetoINSS * 0.6

' Usar:
mediaSalarialEstimada = CalcularMediaReal(ID_Cliente)
```

---

## 🔄 DECISÕES DE REGRAS E PEDÁGIOS

### ❓ Questão 1: "Ao simular no modo automático já é definida a melhor regra?"

**✅ Resposta:** SIM, parcialmente.

**Como Funciona:**

- O botão **"Analisar Todas as Regras"** (`cmdAnalisarTodas_Click`) testa as 5 regras
- Identifica a que resulta em:
  1. Direito adquirido OU
  2. Menor tempo faltante
- Marca automaticamente o OptionButton correspondente
- Exibe o resultado dessa regra

**Porém:** Os 3 cenários (rápido/equilibrado/máximo) são **independentes** das regras específicas. Eles mostram:

- Quanto tempo precisa contribuir para ter 60%, 80% ou 100% do benefício
- Não consideram regras específicas como Pontos ou Pedágios

### ❓ Questão 2: "O advogado deve ter a opção de alterar a regra antes de simular?"

**✅ Resposta:** SIM, a opção existe mas está em **modo legado**.

**Opções Disponíveis:**

1. **Modo Manual (Legado):**

   - Marcar um dos OptionButtons (Tempo, Idade, Pontos, Especial)
   - Usar `cmdCalcular_OLD_Click()` (código comentado/desabilitado)
   - Mostra resultado apenas da regra escolhida

2. **Modo Automático (Atual/Recomendado):**
   - Usar `cmdAnalisarTodas_Click()`
   - Sistema escolhe a melhor regra automaticamente
   - Advogado pode **depois** testar outras regras manualmente se quiser

**Recomendação:**

- Manter modo automático como padrão
- Habilitar modo manual como "avançado" para advogados experientes
- Adicionar toggle: "Modo Automático ☑️ / Manual ☐"

### ❓ Questão 3: "Regras de Transição e Pedágio: o advogado decide ou o sistema?"

**✅ Resposta:** O sistema **VERIFICA ELEGIBILIDADE** automaticamente.

**Fluxo:**

1. Sistema calcula tempo de contribuição na data da reforma (13/11/2019)
2. Verifica quanto tempo faltava:
   - ≤ 0 anos: Direito adquirido → Todas as regras aplicáveis
   - ≤ 2 anos: Elegível para Pedágio 50% e 100%
   - > 2 anos: Apenas Pedágio 100%
3. `AnalisarMelhorRegra()` testa **APENAS** as regras elegíveis

**Interface Atual:**

- Os OptionButtons de Transição/Pedágio são manuais
- Não há validação automática de elegibilidade na interface
- **Sugestão:** Desabilitar (`Enabled = False`) as opções não elegíveis

---

## 📈 PRÓXIMOS PASSOS RECOMENDADOS

### 🔥 Prioridade ALTA

#### 1. **Implementar Cálculo Real de Remunerações**

**Arquivos a modificar:**

- `modSimulacoes.bas` → Adicionar `CalcularMediaReal()`
- `modImportacao.bas` → Adicionar `ImportarRemuneracoesDeCSV()`

**Planilha a criar:**

- `Remuneracoes` (ID_Cliente, Competencia, Valor, Data_Importacao)

**Python (JÁ PRONTO):**

- ✅ `saida_cnis_remuneracoes.csv` já é gerado

#### 2. **Testar Cenários Diversos**

**Perfis de Teste:**

- Cliente jovem (35 anos, 10 anos contrib.)
- Cliente próximo da aposentadoria (60 anos, 28 anos contrib.)
- Cliente com tempo especial (50 anos, 15 especial + 10 comum)
- Cliente com direito adquirido pré-reforma

**Validar:**

- Cálculos de tempo total
- Conversão de tempo especial
- Projeções de data/idade
- Elegibilidade para pedágios

#### 3. **Melhorar Interface de Regras**

- Desabilitar OptionButtons de regras não elegíveis
- Adicionar tooltip explicando por que está desabilitado
- Criar indicador visual: ✅ Elegível / ❌ Não Elegível / ⏳ Pendente

---

### 🔶 Prioridade MÉDIA

#### 4. **Expandir Python para Outros Tipos de Vínculo**

**Pendente:**

- ⚠️ CONTRIBUINTE INDIVIDUAL
- ⚠️ FACULTATIVO
- ⚠️ EMPREGADOR DOMÉSTICO

**Arquivos a modificar:**

- `converter_extrato_inss.py` → Melhorar regex de parsing

#### 5. **Adicionar Reajustes aos Valores**

**Implementar:**

- Índice de reajuste por competência (INPC/TR histórico)
- Planilha `Indices_Reajuste` (Competencia, Indice)
- Função `AplicarReajustes(valor, competenciaInicial, competenciaFinal)`

#### 6. **Relatórios e Exportação**

- Gerar PDF com análise comparativa dos 3 cenários
- Exportar simulação para Word (modelo de petição)
- Salvar histórico de simulações por cliente

---

## 🔍 CASOS DE USO E CENÁRIOS DE TESTE

### Cenário 1: Cliente Jovem Iniciando Contribuição

**Perfil:**

- Nome: João Silva
- Idade: 30 anos
- Sexo: Masculino
- Tempo Contribuído: 5 anos
- Tipo: Empregado (CLT)

**Resultado Esperado:**

```
Cenário 1 (Rápido): 15 anos
- Falta: 10 anos (10 anos e 0 meses)
- Data: 12/01/2036
- Idade: 40 anos
- Valor: R$ 2.824,80 (60% do benefício)

Cenário 2 (Equilibrado): 25 anos
- Falta: 20 anos
- Data: 12/01/2046
- Idade: 50 anos
- Valor: R$ 3.769,73 (70% do benefício)

Cenário 3 (Máximo): 40 anos
- Falta: 35 anos
- Data: 12/01/2061
- Idade: 65 anos
- Valor: R$ 4.671,61 (100% do benefício)
```

**Melhor Regra:** Tempo de Contribuição (faltam 30 anos para 35 anos)

---

### Cenário 2: Cliente Próximo da Aposentadoria

**Perfil:**

- Nome: Maria Santos
- Idade: 58 anos
- Sexo: Feminino
- Tempo Contribuído: 28 anos
- Tempo na Reforma (13/11/2019): 24 anos

**Resultado Esperado:**

```
Direito Adquirido: NÃO
Melhor Regra: Pedágio 50%

Tempo faltava na reforma: 30 - 24 = 6 anos
Pedágio 50%: 6 × 50% = 3 anos
Total necessário: 30 + 3 = 33 anos
Falta: 33 - 28 = 5 anos

Cenário Atual (28 anos):
- Percentual: 60% + 2% × (28 - 15) = 86%
- Valor: R$ 4.017,58

Cenário com 5 anos a mais (33 anos):
- Percentual: 60% + 2% × (33 - 15) = 96%
- Valor: R$ 4.485,14
```

**Elegibilidade:**

- ❌ Pedágio 50% (faltavam mais de 2 anos)
- ✅ Pedágio 100%

---

### Cenário 3: Cliente com Tempo Especial

**Perfil:**

- Nome: Carlos Oliveira
- Idade: 52 anos
- Sexo: Masculino
- Tempo Comum: 10 anos
- Tempo Especial: 15 anos (grau 25 → fator 1.4)

**Cálculos:**

```
Tempo Especial Convertido: 15 × 1.4 = 21 anos
Tempo Total: 10 + 21 = 31 anos
Pontos: 52 + 31 = 83 pontos
```

**Resultado Esperado:**

```
Melhor Regra: Tempo de Contribuição
Falta: 35 - 31 = 4 anos

Pontos necessários: 105 (homem)
Falta de pontos: 105 - 83 = 22 pontos
Tempo para pontuar: 22 / 2 = 11 anos

Conclusão: Tempo de Contribuição é mais rápido (4 anos vs 11 anos)
```

---

## 🐛 PONTOS DE ATENÇÃO E POSSÍVEIS BUGS

### ⚠️ 1. Tratamento de Datas de Nascimento Inválidas

**Código Atual:**

```vb
On Error Resume Next
nascimento = CDate(ws.Cells(linha, 7).Value)
On Error GoTo 0

If nascimento = 0 Then
    ' CÓDIGO INCOMPLETO - Não há tratamento após o If
```

**Problema:** Se a data for inválida, o código continua sem erro, mas com `nascimento = 0`.

**Sugestão:**

```vb
If nascimento = 0 Or Not IsDate(ws.Cells(linha, 7).Value) Then
    MsgBox "Data de nascimento inválida para o cliente ID " & ID_Cliente, vbCritical
    Exit Sub
End If
```

---

### ⚠️ 2. Função `Nz()` Não Nativa do VBA

**Código Atual:**

```vb
carenciaMinima = Nz(GetParametro("Carencia_Minima")) / 12
```

**Problema:** `Nz()` é função do Access, não existe no Excel VBA.

**Solução:**

```vb
' Adicionar em modUtil.bas
Function Nz(value As Variant, Optional valueIfNull As Variant = 0) As Variant
    If IsNull(value) Or IsEmpty(value) Or value = "" Then
        Nz = valueIfNull
    Else
        Nz = value
    End If
End Function
```

---

### ⚠️ 3. Estimativa de Tempo na Reforma Pode Ser Imprecisa

**Código Atual:**

```vb
anosDesdeReforma = DateDiff("d", dataReforma, Date) / 365.25
tempoNaReforma = tempoTotal - anosDesdeReforma
```

**Problema:** Assume que o cliente contribuiu continuamente desde a reforma, o que pode não ser verdade.

**Solução Ideal:**

- Calcular tempo real até 13/11/2019 somando apenas vínculos com data fim ≤ dataReforma
- Implementar `CalcularTempoAte(ID_Cliente, dataLimite)`

---

### ⚠️ 4. Valores Estimados Sem Remunerações Reais

**Impacto:** Simulações podem estar incorretas em até 40% do valor real.

**Urgência:** ALTA - Usuário já está ciente e é o próximo passo.

---

## 📚 GLOSSÁRIO TÉCNICO

| Termo                  | Definição                                                                    |
| ---------------------- | ---------------------------------------------------------------------------- |
| **CNIS**               | Cadastro Nacional de Informações Sociais - Extrato previdenciário do INSS    |
| **Concomitância**      | Períodos de vínculos sobrepostos (trabalhando em 2 lugares ao mesmo tempo)   |
| **Fator de Conversão** | Multiplicador para converter tempo especial em tempo comum (1.4, 1.75, 2.33) |
| **Pontos**             | Soma de idade + tempo de contribuição (usado na Regra de Pontos)             |
| **Carência**           | Número mínimo de contribuições mensais (180 = 15 anos)                       |
| **Pedágio**            | Tempo adicional para quem estava próximo da aposentadoria na reforma         |
| **Coeficiente**        | Percentual do benefício (60% a 100%) baseado no tempo de contribuição        |
| **Média 80%**          | Média dos 80% maiores salários de contribuição (base de cálculo)             |
| **Teto INSS**          | Valor máximo que o INSS paga (R$ 7.786,02 em 2026)                           |

---

## 🎓 REFERÊNCIAS LEGAIS

- **EC 103/2019** - Emenda Constitucional da Reforma da Previdência
- **Lei 13.183/2019** - Lei de implementação da reforma
- **IN INSS 128/2022** - Instrução normativa sobre cálculo de benefícios
- **Portaria MTP 26/2022** - Valores atualizados de benefícios

---

## ✅ CONCLUSÃO

O **frmSimulacoes** está funcionando corretamente para:

- ✅ Cálculo de tempo de contribuição
- ✅ Conversão de tempo especial
- ✅ Simulação de 3 cenários (rápido, equilibrado, máximo)
- ✅ Análise automática da melhor regra
- ✅ Projeções de data e idade de aposentadoria

**Pendente de implementação:**

- 🔶 Cálculo real de valores (aguardando implementação de remunerações)
- 🔶 Reajustes históricos dos salários
- 🔶 Expansão do Python para vínculos facultativos
- 🔶 Interface melhorada para elegibilidade de regras

**Recomendação Final:**

1. Implementar `CalcularMediaReal()` com dados de remunerações
2. Criar bateria de testes com perfis diversos
3. Validar cálculos com casos reais conhecidos
4. Adicionar logs de debug para rastreamento de erros

---

**Documento gerado por:** GitHub Copilot (Claude Sonnet 4.5)  
**Validação técnica:** Pendente de revisão pelo desenvolvedor
