# 📊 EXEMPLOS DE CÁLCULO - Validação com Advogado

**Data:** 22 de janeiro de 2026  
**Objetivo:** Demonstrar cálculos para validação do parceiro jurídico

---

## 🧮 PARTE 1: Conversão de Tempo Especial

### Exemplo 1: Tempo Especial 15 anos (Mulher)

**Dados do Vínculo:**
- Período: 01/01/2010 a 31/12/2019 (10 anos exatos)
- Tipo: Empregado
- Especial: Sim
- Grau: 15 anos
- Sexo do Cliente: Feminino

**Cálculo pelo Sistema:**
```
Tempo Real = 10 anos
Fator de Conversão (15 anos, Mulher) = 2.00
Tempo Convertido = 10 × 2.00 = 20 anos
```

**✅ Validação:** 10 anos de trabalho em condições especiais (grau 15) conta como 20 anos de tempo comum.

---

### Exemplo 2: Tempo Especial 20 anos (Homem)

**Dados do Vínculo:**
- Período: 01/01/2005 a 31/12/2019 (15 anos exatos)
- Tipo: Empregado
- Especial: Sim
- Grau: 20 anos
- Sexo do Cliente: Masculino

**Cálculo pelo Sistema:**
```
Tempo Real = 15 anos
Fator de Conversão (20 anos, Homem) = 1.75
Tempo Convertido = 15 × 1.75 = 26.25 anos (26 anos e 3 meses)
```

**✅ Validação:** 15 anos de trabalho especial (grau 20) conta como 26 anos e 3 meses de tempo comum.

---

### Exemplo 3: Tempo Especial 25 anos (Mulher)

**Dados do Vínculo:**
- Período: 01/03/2000 a 28/02/2020 (20 anos exatos)
- Tipo: Empregado
- Especial: Sim
- Grau: 25 anos
- Sexo do Cliente: Feminino

**Cálculo pelo Sistema:**
```
Tempo Real = 20 anos
Fator de Conversão (25 anos, Mulher) = 1.20
Tempo Convertido = 20 × 1.20 = 24 anos
```

**✅ Validação:** 20 anos de trabalho especial (grau 25) conta como 24 anos de tempo comum.

---

## 📋 TABELA COMPLETA DE FATORES

| Grau | Sexo | Fator | Exemplo: 10 anos reais | Sistema Usa |
|------|------|-------|------------------------|-------------|
| 15   | H    | 2.33  | 10 × 2.33 = 23.3 anos  | ⚠️ **VALIDAR** |
| 15   | M    | 2.00  | 10 × 2.00 = 20.0 anos  | ✅ Confirmado |
| 20   | H    | 1.75  | 10 × 1.75 = 17.5 anos  | ✅ Confirmado |
| 20   | M    | 1.50  | 10 × 1.50 = 15.0 anos  | ✅ Confirmado |
| 25   | H    | 1.40  | 10 × 1.40 = 14.0 anos  | ✅ Confirmado |
| 25   | M    | 1.20  | 10 × 1.20 = 12.0 anos  | ✅ Confirmado |

**⚠️ PENDENTE:** Advogado não informou fator para **15 anos Homem**. Sistema usa 2.33.

---

## 🧮 PARTE 2: Cálculo de Aposentadoria por Tempo de Contribuição

### Exemplo 4: Homem - Regra de Tempo (Direito Adquirido)

**Dados do Cliente:**
- Nome: João Silva
- Sexo: Masculino
- Data Nascimento: 15/03/1965
- Data Simulação: 13/11/2019 (antes da reforma)

**Vínculos:**
1. 01/01/1985 a 31/12/2004 (20 anos) - Comum
2. 01/01/2005 a 13/11/2019 (14.87 anos) - Especial Grau 25 (Homem)

**Cálculo:**
```
Tempo Comum = 20 anos
Tempo Especial Convertido = 14.87 × 1.40 = 20.82 anos
Tempo Total = 20 + 20.82 = 40.82 anos
```

**Verificação da Regra:**
- Requisito: 35 anos de contribuição (Homem)
- Cliente tem: 40.82 anos ✅
- Data da regra: 13/11/2019 (antes da reforma) ✅

**✅ RESULTADO:** Cliente tem direito adquirido à aposentadoria por tempo de contribuição.

---

## 🧮 PARTE 3: Cálculo de Aposentadoria por Pontos

### Exemplo 5: Mulher - Regra de Pontos

**Dados do Cliente:**
- Nome: Maria Santos
- Sexo: Feminino
- Data Nascimento: 10/05/1970
- Data Simulação: 22/01/2026
- Idade atual: 55 anos e 8 meses

**Vínculos:**
1. 01/01/1990 a 31/12/2010 (21 anos) - Comum
2. 01/01/2011 a 22/01/2026 (15.06 anos) - Especial Grau 20 (Mulher)

**Cálculo:**
```
Tempo Comum = 21 anos
Tempo Especial Convertido = 15.06 × 1.50 = 22.59 anos
Tempo Total = 21 + 22.59 = 43.59 anos

Pontos = Idade + Tempo
Pontos = 55.67 + 43.59 = 99.26 pontos
```

**Verificação da Regra:**
- Requisito Mulher: 100 pontos + mínimo 30 anos
- Cliente tem: 99.26 pontos ❌
- Tempo: 43.59 anos ✅

**❌ RESULTADO:** Faltam 0.74 pontos (aproximadamente 9 meses).

---

## 🧮 PARTE 4: Correção Monetária - INPC vs SELIC

### Exemplo 6: Comparação de Índices

**Salário de Janeiro/1995:** R$ 1.000,00  
**Correção até Janeiro/2026:**

#### INPC (Série 188):
```
Fator INPC (01/1995): 13.52
Valor Corrigido = R$ 1.000 × 13.52 = R$ 13.520,00
```

#### SELIC (Série 1178):
```
Fator SELIC (01/1995): 49.468.933.101.497.159.254.016,00
Valor Corrigido = R$ 1.000 × ... = número astronômico
```

**⚠️ IMPORTANTE:** 
- SELIC acumula juros (crescimento exponencial)
- INPC corrige apenas inflação (crescimento linear)
- Para revisões previdenciárias, usar **INPC** (correção legal)
- SELIC só é usada para mora/judiciais (não é caso do sistema)

**✅ RECOMENDAÇÃO:** Sistema deve usar **INPC**, não SELIC, para correção de salários.

---

## ❓ QUESTÕES PARA O ADVOGADO

### 1. Fator de Conversão 15 anos Homem
**Sistema usa:** 2.33  
**Advogado informou:** ❓ (não mencionou)  
**❓ CONFIRMAR:** Fator 2.33 está correto?

### 2. Índice de Correção
**Sistema tem:** INPC (série 188)  
**Script criado:** SELIC (série 1178)  
**❓ CONFIRMAR:** Qual índice usar para correção de salários?

### 3. Aposentadoria Especial Pura
**Situação:** Cliente tem 25 anos PUROS em atividade especial grau 15  
**❓ CONFIRMAR:** Sistema deve permitir aposentadoria especial (sem conversão)?

### 4. Validação de Cálculos
**Pedido:** Fornecer **1 caso real** completo para validação:
- Sexo, idade, data nascimento
- Lista de vínculos (datas, tipo, especial?, grau?)
- Resultado esperado (regra + tempo + benefício)

---

## 📝 COMO USAR ESTE DOCUMENTO

1. **Imprimir ou enviar ao advogado**
2. **Pedir validação** dos exemplos 1-6
3. **Solicitar caso real** para teste completo
4. **Confirmar decisões** sobre as 4 questões pendentes

---

**Última atualização:** 22/01/2026  
**Responsável:** Equipe de Desenvolvimento Sist_Prev
