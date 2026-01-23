# 📊 MAPEAMENTO COMPLETO - Planilha INSS Simulação

**Data da Análise:** 22 de janeiro de 2026
**Arquivo:** inss_simulacao.xlsm
**Total de Abas:** 6

---

## 📑 ÍNDICE DE ABAS

1. [1.Dados Gerais do Contribuinte](#1.dados-gerais-do-contribuinte) - 20 células
2. [2.Historico de contribuições](#2.historico-de-contribuições) - 23 células
3. [3.Indice SELIC acumulado](#3.indice-selic-acumulado) - 1426 células
4. [4.Calculo da media salarial](#4.calculo-da-media-salarial) - 21 células
5. [5.Calculo do coeficiente](#5.calculo-do-coeficiente) - 6 células
6. [6.Resultado](#6.resultado) - 8 células

---

## 1.Dados Gerais do Contribuinte

**Dimensões:** 14 linhas × 3 colunas

### 📌 Estrutura de Colunas

| Coluna | Cabeçalho | Tipo Detectado |
|--------|-----------|----------------|
| A | - | - |
| B | - | Texto |
| C | - | Texto |

### 📊 Amostra de Dados (5 primeiras linhas)

```
Linha 1:  |  | 
Linha 2:  | Campo | Valor
Linha 3:  | Nome do contribuinte | João Carlos Eduardo Figueir...
Linha 4:  | Data de nascimento | 1973-03-22 00:00:00
Linha 5:  | Sexo | Masculino
```

### 🔧 Fórmulas Principais

*Nenhuma fórmula encontrada*

---

## 2.Historico de contribuições

**Dimensões:** 11 linhas × 5 colunas

### 📌 Estrutura de Colunas

| Coluna | Cabeçalho | Tipo Detectado |
|--------|-----------|----------------|
| A | - | - |
| B | - | Texto |
| C | - | Texto |
| D | - | - |
| E | - | - |

### 📊 Amostra de Dados (5 primeiras linhas)

```
Linha 1:  |  |  |  | 
Linha 2:  | Coluna | Descrição |  | 
Linha 3:  | Inicio |  |  | Data de início do vínculo
Linha 4:  | Fim |  |  | Data de fim do vínculo
Linha 5:  | Tipo de atividade |  |  | Comum / Especial 25 / Espec...
```

### 🔧 Fórmulas Principais

```excel
C7: =C4-C3+1
C8: =C7 * C6
C11: =C9*C10
```

---

## 3.Indice SELIC acumulado

**Dimensões:** 500 linhas × 7 colunas

### 📌 Estrutura de Colunas

| Coluna | Cabeçalho | Tipo Detectado |
|--------|-----------|----------------|
| A | - | - |
| B | Mês / Ano | Outro |
| C | Índice SELIC acumulado | Número |
| D | Fonte | Texto |
| E | - | - |
| F | - | - |
| G | - | - |

### 📊 Amostra de Dados (5 primeiras linhas)

```
Linha 1:  | Mês / Ano | Índice SELIC acumulado | Fonte |  |  | 
Linha 2:  | 1986-01-08 00:00:00 | 2.57 | BCB API |  |  | 
Linha 3:  | 1986-01-09 00:00:00 | 2.94 | BCB API |  |  | 
Linha 4:  | 1986-01-10 00:00:00 | 1.96 | BCB API |  |  | Atualizado em: 22/01/2026 1...
Linha 5:  | 1986-01-11 00:00:00 | 2.37 | BCB API |  |  | 
```

### 🔧 Fórmulas Principais

*Nenhuma fórmula encontrada*

---

## 4.Calculo da media salarial

**Dimensões:** 302 linhas × 8 colunas

### 📌 Estrutura de Colunas

| Coluna | Cabeçalho | Tipo Detectado |
|--------|-----------|----------------|
| A | - | - |
| B | Data | Outro |
| C | Salario | Número |
| D | Índice SELIC | Fórmula |
| E | Salario corrigido | - |
| F | - | - |
| G | Média = Soma dos Salarios / Quantidade de Meses | Texto |
| H | - | - |

### 📊 Amostra de Dados (5 primeiras linhas)

```
Linha 1:  | Data | Salario | Índice SELIC | Salario corrigido |  | Média = Soma dos Salarios /... | 
Linha 2:  | 1994-07-01 00:00:00 | 1235.56 | =VLOOKUP(B2,'3.Indice SELIC... |  |  | Resumo da média salarial | 
Linha 3:  | 1994-08-01 00:00:00 | 1250 |  |  |  | Quantidade Meses | =B302
Linha 4:  |  |  |  |  |  |  | 
Linha 5:  |  |  |  |  |  | Somatória de Salários corri... | =C302
```

### 🔧 Fórmulas Principais

```excel
D2: =VLOOKUP(B2,'3.Indice SELIC acumulado'!B2:C299,2,FALSE)
H3: =B302
H5: =C302
H7: =H5/H3
```

---

## 5.Calculo do coeficiente

**Dimensões:** 5 linhas × 3 colunas

### 📌 Estrutura de Colunas

| Coluna | Cabeçalho | Tipo Detectado |
|--------|-----------|----------------|
| A | - | - |
| B | - | Texto |
| C | - | Texto |

### 📊 Amostra de Dados (5 primeiras linhas)

```
Linha 1:  |  | 
Linha 2:  | Campo | Valor
Linha 3:  | Anos totais de contribuição | 
Linha 4:  | Anos acima dos mínimo      ... | 
Linha 5:  | Coeficiente | =60% + 2% * C4
```

### 🔧 Fórmulas Principais

```excel
C5: =60% + 2% * C4
```

---

## 6.Resultado

**Dimensões:** 6 linhas × 3 colunas

### 📌 Estrutura de Colunas

| Coluna | Cabeçalho | Tipo Detectado |
|--------|-----------|----------------|
| A | - | - |
| B | - | Texto |
| C | - | Texto |

### 📊 Amostra de Dados (5 primeiras linhas)

```
Linha 1:  |  | 
Linha 2:  | Descrição | Valor
Linha 3:  | Coluna1 | Coluna2
Linha 4:  | Média salarial corrigida | 
Linha 5:  | Coeficiente | 
```

### 🔧 Fórmulas Principais

```excel
C6: =C4*C5
```

---

