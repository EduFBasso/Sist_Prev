# 📋 GUIA COMPLETO: Regras de Aposentadoria - Sistema Previdenciário

**Data:** 21 de janeiro de 2026  
**Sistema:** Sist_Prev v1.1  
**Público-alvo:** Advogados e desenvolvedores

---

## 🎯 VISÃO GERAL DO SISTEMA

O sistema trabalha com **DOIS NÍVEIS DE REGRAS**:

### 1️⃣ Regras por VÍNCULO (características do período)
- Definidas no cadastro de vínculos (`frmVinculos`)
- Afetam o cálculo de **tempo de contribuição**
- **Especial (15/20/25 anos)** → conversão para tempo comum
- Flags: Rural, Militar, Exterior, etc.

### 2️⃣ Regras por SIMULAÇÃO (regras de aposentadoria)
- Aplicadas no motor de cálculo (`modSimulacoes`)
- Determinam **qual tipo de aposentadoria** o cliente pode obter
- 5 regras principais + cenários de tempo

---

## 📚 AS 5 REGRAS DE APOSENTADORIA

### ✅ 1. TEMPO DE CONTRIBUIÇÃO (Direito Adquirido)

**Quando usar:**  
Clientes que **completaram o tempo ANTES da reforma** (13/11/2019)

**Requisitos:**
- **Homem:** 35 anos de contribuição
- **Mulher:** 30 anos de contribuição
- **Não há requisito de idade mínima**

**Vantagem:** Pode aposentar a qualquer idade  
**Desvantagem:** Só vale para quem tem direito adquirido

**Código:** `RegraTempoContribuicao()` em [modSimulacoes.bas](../Modulos/modSimulacoes.bas)

---

### ✅ 2. IDADE MÍNIMA (Regra Permanente Pós-Reforma)

**Quando usar:**  
Regra padrão atual do INSS (válida para todos)

**Requisitos:**
- **Homem:** 65 anos de idade + 15 anos de contribuição (carência)
- **Mulher:** 62 anos de idade + 15 anos de contribuição (carência)

**Vantagem:** Menor tempo de contribuição necessário (apenas 15 anos)  
**Desvantagem:** Exige idade mínima alta

**Código:** `RegraIdade()` em [modSimulacoes.bas](../Modulos/modSimulacoes.bas)

**Observação:** A idade mínima da mulher está em transição progressiva:
- 2019-2023: 60 anos
- 2024-2026: 61 anos  
- 2027+: 62 anos

---

### ✅ 3. REGRA DE PONTOS (Transição)

**Quando usar:**  
Para quem não tem direito adquirido mas quer se aposentar antes da idade mínima

**Requisitos:**
- **Pontos = Idade + Tempo de Contribuição**
- **Homem:** 105 pontos + mínimo 35 anos de contribuição
- **Mulher:** 100 pontos + mínimo 30 anos de contribuição

**Exemplo prático:**
- Homem com 60 anos + 45 anos de contribuição = 105 pontos ✅
- Mulher com 55 anos + 45 anos de contribuição = 100 pontos ✅

**Vantagem:** Permite aposentadoria antes da idade mínima se tiver tempo suficiente  
**Desvantagem:** Exige muito tempo de contribuição

**Código:** `RegraPontos()` em [modSimulacoes.bas](../Modulos/modSimulacoes.bas)

**Observação:** A pontuação aumenta 1 ponto por ano até atingir o máximo

---

### ✅ 4. PEDÁGIO 50% (Transição)

**Quando usar:**  
**SOMENTE** para quem estava a **ATÉ 2 ANOS** de se aposentar na data da reforma (13/11/2019)

**Requisitos:**
- Faltavam **até 2 anos** para completar 35H/30M em 13/11/2019
- Precisa cumprir: **tempo necessário + 50% do que faltava**

**Exemplo prático:**
- Homem tinha 34 anos em 13/11/2019 (faltava 1 ano)
- Pedágio: 1 ano × 50% = 6 meses
- **Total necessário:** 35 anos + 6 meses = 35,5 anos

**Vantagem:** Pedágio menor (50%), sem exigência de idade mínima  
**Desvantagem:** Só para quem estava muito perto na reforma

**Código:** `RegraPedagio50()` em [modSimulacoes.bas](../Modulos/modSimulacoes.bas)

---

### ✅ 5. PEDÁGIO 100% (Transição)

**Quando usar:**  
Para quem estava **a MAIS de 2 anos** de se aposentar na reforma, mas tinha pelo menos 28H/25M anos

**Requisitos:**
- Homem: tinha pelo menos 28 anos em 13/11/2019
- Mulher: tinha pelo menos 25 anos em 13/11/2019
- Precisa cumprir: **tempo necessário + 100% do que faltava**
- **Idade mínima:** 60H / 57M

**Exemplo prático:**
- Mulher tinha 27 anos em 13/11/2019 (faltavam 3 anos)
- Pedágio: 3 anos × 100% = 3 anos
- **Total necessário:** 30 anos + 3 anos = 33 anos
- **Idade mínima:** 57 anos

**Vantagem:** Permite aposentadoria mesmo estando longe em 2019  
**Desvantagem:** Pedágio maior (100%) + exige idade mínima

**Código:** `RegraPedagio100()` em [modSimulacoes.bas](../Modulos/modSimulacoes.bas)

---

## 🏗️ APOSENTADORIA ESPECIAL (Implementação Pendente)

### ⚠️ STATUS ATUAL

A Aposentadoria Especial está **PARCIALMENTE IMPLEMENTADA**:

✅ **Implementado:**
- Marcação de vínculos especiais em `frmVinculos`
- Grau de especialidade (15/20/25 anos)
- Conversão de tempo especial para comum
- Fatores diferenciados por sexo

❌ **Pendente:**
- Regra de aposentadoria especial pura (sem conversão)
- Cálculo de valor do benefício para especial
- Integração com frmSimulacoes

---

### 📐 COMO FUNCIONA TEMPO ESPECIAL

**Atividades especiais** (insalubres, periculosas ou de risco) podem ser:

1. **CONVERTIDAS para tempo comum** (usando fatores de multiplicação)
2. **USADAS para aposentadoria especial** (sem conversão, tempo puro)

---

### 🔄 CONVERSÃO DE TEMPO ESPECIAL → COMUM

O sistema aplica fatores diferenciados por **sexo** e **grau**:

| Grau | Atividade | Tempo Especial | Fator Homem | Fator Mulher |
|------|-----------|----------------|-------------|--------------|
| 15   | Alta insalubridade | 15 anos | 2.33 | 2.00 |
| 20   | Média insalubridade | 20 anos | 1.75 | 1.50 |
| 25   | Baixa insalubridade | 25 anos | 1.40 | 1.20 |

**Exemplo prático:**
- Mulher trabalhou 10 anos em atividade especial grau 15
- Conversão: 10 anos × 2.00 = **20 anos de tempo comum**
- Esses 20 anos contam para qualquer regra (Pontos, Idade, etc.)

**Código:** `CalcularTempoEspecial()` em [modSimulacoes.bas](../Modulos/modSimulacoes.bas)

**Parâmetros configuráveis em Config_Regras:**
- `Conversao_Especial_15_H` = 2.33
- `Conversao_Especial_15_M` = 2.00
- `Conversao_Especial_20_H` = 1.75
- `Conversao_Especial_20_M` = 1.50
- `Conversao_Especial_25_H` = 1.40
- `Conversao_Especial_25_M` = 1.20

---

### 🎯 APOSENTADORIA ESPECIAL PURA (Regra 6 - Pendente)

**Quando deveria usar:**
Quando o cliente quer se aposentar PELA atividade especial (sem converter)

**Requisitos (a implementar):**
- **15 anos:** Atividades de alto risco (mineração subterrânea, amianto)
- **20 anos:** Atividades de médio risco (exposição a agentes químicos)
- **25 anos:** Atividades de baixo risco (ruído acima de 85dB)
- **Não há requisito de idade mínima**

**Vantagem:** Aposentadoria em menos tempo (15/20/25 anos)  
**Desvantagem:** 
- Valor do benefício pode ser menor
- Precisa comprovar atividade especial por TODO o período
- Não pode trabalhar em atividade especial após se aposentar

**Status:** Botão `optEspecial` existe no formulário mas mostra mensagem:
> "AVISO: A regra 'Aposentadoria Especial' foi identificada, mas ainda não está completamente implementada."

---

## 🔍 FLUXO DE DECISÃO AUTOMÁTICA

O sistema possui uma função `AnalisarMelhorRegra()` que:

### Ordem de verificação:

1. **Tempo de Contribuição** → Se já tem direito adquirido, retorna imediatamente
2. **Idade** → Se já atende idade + carência, retorna
3. **Pontos** → Se já atinge pontuação + tempo, retorna
4. **Pedágio 50%** → Se elegível e atende requisitos, retorna
5. **Pedágio 100%** → Se elegível e atende requisitos, retorna
6. **Melhor opção** → Retorna a regra com menor tempo faltante

**Código:** `AnalisarMelhorRegra()` em [modSimulacoes.bas](../Modulos/modSimulacoes.bas#L790)

---

## 🎬 CENÁRIOS DE SIMULAÇÃO

Além das regras, o sistema calcula **3 CENÁRIOS DE TEMPO**:

### 📊 Cenário 1: APOSENTADORIA RÁPIDA
- **Tempo:** 15 anos (mínimo legal)
- **Percentual:** 60% da média salarial
- **Quando:** Cliente quer se aposentar o mais rápido possível
- **Valor:** Menor benefício

### 📊 Cenário 2: APOSENTADORIA EQUILIBRADA (RECOMENDADA)
- **Tempo:** 25 anos
- **Percentual:** 80% da média salarial (60% + 10 anos × 2%)
- **Quando:** Balanço entre tempo de espera e valor do benefício
- **Valor:** Intermediário

### 📊 Cenário 3: APOSENTADORIA MÁXIMA
- **Tempo:** 40 anos
- **Percentual:** 100% da média salarial (60% + 25 anos × 2%)
- **Quando:** Cliente quer benefício integral
- **Valor:** Máximo

**Cálculo do percentual:**
- Base: 60%
- Adicional: +2% por ano acima de 15 anos (mulher) ou 20 anos (homem)
- Máximo: 100%

**Código:** 
- `CalcularCenarioRapido()` em [modSimulacoes.bas](../Modulos/modSimulacoes.bas#L931)
- `CalcularCenarioEquilibrado()` em [modSimulacoes.bas](../Modulos/modSimulacoes.bas)
- `CalcularCenarioMaximo()` em [modSimulacoes.bas](../Modulos/modSimulacoes.bas)

---

## 🎛️ MODO AUTOMÁTICO vs MANUAL

### 🟢 MODO AUTOMÁTICO (Padrão)

**Comportamento:**
1. Usuário clica em "Calcular"
2. Sistema analisa automaticamente a melhor regra
3. Pré-seleciona o OptionButton correspondente
4. Desabilita os outros OptionButtons (foco visual)
5. Exibe resultado com a regra escolhida

**Vantagens:**
- Rápido e intuitivo
- Sistema decide a melhor opção
- Menos chance de erro do usuário

**Quando usar:** Análise inicial, casos simples

---

### 🟠 MODO MANUAL

**Comportamento:**
1. Usuário clica no botão Toggle (muda para "Manual")
2. Todos os OptionButtons são **habilitados** e **desmarcados**
3. Usuário PRECISA selecionar uma regra manualmente
4. Sistema calcula baseado na regra escolhida

**Vantagens:**
- Controle total do advogado
- Permite testar regras específicas
- Útil para análise comparativa

**Quando usar:** Análise detalhada, comparações, casos complexos

**Correção implementada:**
- Agora ao mudar para Manual, os botões são habilitados E desmarcados
- Força o usuário a fazer uma escolha consciente
- Evita confusão com seleção anterior do modo Auto

---

## ⚙️ PARÂMETROS CONFIGURÁVEIS

Os parâmetros estão em `Config_Regras` (Planilha) e podem ser alterados:

### Requisitos Gerais
| Parâmetro | Valor Padrão | Descrição |
|-----------|--------------|-----------|
| `Carencia_Minima` | 180 | Carência mínima em meses (15 anos) |
| `Data_Reforma` | 13/11/2019 | Data da Reforma da Previdência |

### Regra de Idade
| Parâmetro | Valor Padrão | Descrição |
|-----------|--------------|-----------|
| `Idade_Minima_Homem` | 65 | Idade mínima para homens |
| `Idade_Minima_Mulher` | 62 | Idade mínima para mulheres |

### Regra de Pontos
| Parâmetro | Valor Padrão | Descrição |
|-----------|--------------|-----------|
| `Pontos_Homem` | 105 | Pontuação necessária para homens |
| `Pontos_Mulher` | 100 | Pontuação necessária para mulheres |

### Conversão Especial
| Parâmetro | Valor Padrão | Descrição |
|-----------|--------------|-----------|
| `Conversao_Especial_15_H` | 2.33 | Fator de conversão 15 anos (homem) |
| `Conversao_Especial_15_M` | 2.00 | Fator de conversão 15 anos (mulher) |
| `Conversao_Especial_20_H` | 1.75 | Fator de conversão 20 anos (homem) |
| `Conversao_Especial_20_M` | 1.50 | Fator de conversão 20 anos (mulher) |
| `Conversao_Especial_25_H` | 1.40 | Fator de conversão 25 anos (homem) |
| `Conversao_Especial_25_M` | 1.20 | Fator de conversão 25 anos (mulher) |

**Como alterar:** 
1. Abra a planilha `Config_Regras`
2. Localize o parâmetro
3. Altere o valor na coluna correspondente
4. O sistema usa automaticamente o novo valor

---

## 🔧 CHECKLIST PARA DEFINIR COM SEU PARCEIRO JURÍDICO

### ✅ Questões sobre Tempo Especial

1. **Conversão automática ou escolha?**
   - [ ] Sistema sempre converte especial → comum (atual)
   - [ ] Usuário escolhe: converter OU aposentar especial
   - [ ] Automático: sistema decide o melhor caminho

2. **Quando permitir aposentadoria especial pura?**
   - [ ] Cliente com 15/20/25 anos só de atividade especial
   - [ ] Cliente com período misto (especial + comum)
   - [ ] Nunca (sempre converter)

3. **Fatores de conversão estão corretos?**
   - [ ] Sim, usar os atuais (2.33, 1.75, 1.40 para homens)
   - [ ] Não, precisa ajustar (especificar novos valores)
   - [ ] Depende da data do vínculo (antes/depois da reforma)

---

### ✅ Questões sobre Correção Monetária

4. **Qual índice usar para corrigir salários históricos?**
   - [ ] INPC (atual no sistema)
   - [ ] SELIC
   - [ ] Depende: INPC até X, SELIC depois
   - [ ] Usuário escolhe por simulação

5. **A correção é para:**
   - [ ] Calcular média dos salários para benefício (implementado)
   - [ ] Calcular atrasados de pagamento (diferente)
   - [ ] Ambos

---

### ✅ Questões sobre Carência

6. **Como contar carência?**
   - [ ] Por meses de contribuição (competências)
   - [ ] Por tempo corrido (períodos de vínculo)
   - [ ] Contagem diferente por tipo de segurado

7. **Carência mínima de 15 anos está correta?**
   - [ ] Sim para aposentadoria por idade
   - [ ] Não, precisa ajustar
   - [ ] Depende da regra aplicada

---

### ✅ Questões sobre Sexo/Gênero

8. **Como tratar campo Sexo no cadastro?**
   - [ ] Sempre "M" ou "F" (padronizar entrada)
   - [ ] Aceitar "Masculino"/"Feminino" e normalizar
   - [ ] Permitir outros valores (como tratar nas regras?)

---

### ✅ Questões sobre Pedágios

9. **Elegibilidade para Pedágio 50%:**
   - [ ] Faltavam até 2 anos em 13/11/2019 (atual)
   - [ ] Outro critério

10. **Elegibilidade para Pedágio 100%:**
    - [ ] Tinha pelo menos 28H/25M anos em 13/11/2019 (atual)
    - [ ] Outro critério

---

### ✅ Questões sobre Cálculo de Benefício

11. **Fórmula de cálculo do valor está correta?**
    - [ ] Sim: 60% + 2% por ano acima de 15/20 anos
    - [ ] Não, precisa ajustar
    - [ ] Depende da regra aplicada

12. **Teto do INSS:**
    - [ ] Sistema já aplica o teto? (verificar)
    - [ ] Usuário deve aplicar manualmente depois
    - [ ] Não aplicar (valor bruto)

---

## 🎓 RESUMO PARA DESENVOLVEDORES

### Onde está cada coisa:

**Vínculos (características dos períodos):**
- UI: `frmVinculos.bas`
- Dados: Planilha `Vinculos` (colunas 6-7: Especial + Grau)
- Cálculo: `CalcularTempoEspecial()` em `modSimulacoes.bas`

**Regras de Aposentadoria:**
- UI: `frmSimulacoes.bas` (OptionButtons + botão Calcular)
- Motor: `modSimulacoes.bas` (funções RegraXXX)
- Decisão: `AnalisarMelhorRegra()` em `modSimulacoes.bas`

**Cenários:**
- Cálculo: `CalcularCenarioRapido/Equilibrado/Maximo()` em `modSimulacoes.bas`
- Exibição: `GerarTextoComparativoCenarios()` em `frmSimulacoes.bas`

**Parâmetros:**
- Armazenamento: Planilha `Config_Regras`
- Acesso: `GetParametro()` e `SetParametro()` em `modUtil.bas`

---

## 🐛 BUG CORRIGIDO

**Problema:** Ao clicar em Toggle para modo MANUAL, a mensagem pedia para escolher outra regra, mas os OptionButtons permaneciam desabilitados.

**Causa:** A função `cmdCalcular_Click()` desabilitava todos os botões no modo AUTO e habilitava apenas o selecionado. Ao mudar para MANUAL, `AtualizarEstadoBotoes()` habilitava os botões, mas eles mantinham a última seleção do modo AUTO.

**Solução:** Modificada `AtualizarEstadoBotoes()` para:
1. Habilitar todos os botões no modo MANUAL
2. **DESMARCAR todos os botões** ao mudar para MANUAL
3. Forçar escolha consciente do usuário

**Arquivo alterado:** [frmSimulacoes.bas](../Formularios/frmSimulacoes.bas#L43-L75)

---

## 📞 PRÓXIMOS PASSOS

### Implementações pendentes:

1. **Aposentadoria Especial completa**
   - Definir quando usar (com parceiro jurídico)
   - Implementar `RegraAposentadoriaEspecial()`
   - Integrar com `AnalisarMelhorRegra()`
   - Calcular valor do benefício específico

2. **Índice SELIC**
   - Implementar coleta de dados SELIC
   - Criar script Python de atualização
   - Adicionar opção de escolha no formulário
   - Aplicar no cálculo de média

3. **Testes automatizados**
   - Criar casos de teste para cada regra
   - Validar cálculos com exemplos reais
   - Comparar com cálculos do INSS

---

**Documento criado em:** 21/01/2026  
**Última atualização:** 21/01/2026  
**Versão:** 1.0
