# Atualização: Extração de Vínculos Facultativos e Contribuições

## 🆕 NOVOS TIPOS DE VÍNCULOS A EXTRAIR

### 1. Vínculos de Recolhimento/Facultativo

**Exemplo do CNIS:**

```
Seq. 11
NIT: 125.37781.66-1
Origem do Vínculo: RECOLHIMENTO
Tipo de Filiado: Facultativo
Data Início: 01/09/2019
Data Fim: 31/10/2019
Indicadores: IREC-INDPEND
```

**Características:**

- Não tem empresa (CNPJ)
- Origem: RECOLHIMENTO ou CONTRIBUINTE INDIVIDUAL
- Tipo: Facultativo, Contribuinte Individual, etc.
- Pode ter indicadores específicos

### 2. Tabela de Contribuições (sem empresa)

**Exemplo do CNIS:**

```
Frame: "Contribuições"

Competência | Data Pgto | Contribuição | Salário Contrib | Indicadores
10/2024     | 07/10/2024| 282,40       | 1.412,00       | PREC-FACULTCONC
11/2024     | 18/11/2024| 282,40       | 1.412,00       | PREC-FACULTCONC
12/2024     | (vazio)   | (vazio)      | (vazio)        | PREC-FACULTCONC
01/2025     | 22/01/2025| 303,60       | 1.518,00       | PREC-FACULTCONC
```

**Características:**

- Tabela com 5 colunas
- Pode ter linhas em branco no meio
- Cada linha "visual" pode ter dados duplicados (2 registros por linha)
- Valores de contribuição e salário

## 🔧 ATUALIZAÇÃO NECESSÁRIA NO SCRIPT

Vou atualizar o `converter_extrato_inss.py` para:

1. **Detectar vínculos sem empresa** (Facultativo, Contribuinte Individual, Recolhimento)
2. **Extrair tabelas de contribuições** específicas desses vínculos
3. **Gerar CSV adicional** com esses vínculos e contribuições

## 📊 ARQUIVOS QUE SERÃO GERADOS

Após atualização, teremos:

- `saida_cnis_vinculos_estruturado.csv` - Vínculos com empresa (como antes)
- `saida_cnis_vinculos_sem_empresa.csv` - 🆕 Vínculos facultativos/recolhimento
- `saida_cnis_contribuicoes_individuais.csv` - 🆕 Contribuições mensais detalhadas
- `saida_cnis_remuneracoes.csv` - Remunerações de vínculos com empresa (como antes)

## 🎯 IMPACTO NO CASO DO JOÃO CARLOS

Com os novos dados extraídos, ele terá:

### Vínculos Já Capturados (1995-2008):

~13-15 anos

### Novos Vínculos a Capturar:

- **Seq. 11:** Facultativo (09/2019 - 10/2019) = 2 meses
- **Seq. 12:** Facultativo (10/2024 - 01/2025) = 4 meses

### Total Estimado:

~13-15 anos + 6 meses = **~13,5-15,5 anos**

Isso ainda não é suficiente para aposentadoria, mas:

- Reduz o tempo faltante
- Mostra contribuições recentes (ativas)
- Demonstra intenção de continuar contribuindo

## ⚠️ OBSERVAÇÃO

Antes de atualizar o script, **execute primeiro a simulação com os dados atuais** para ter uma baseline. Depois atualizamos e comparamos os resultados!

---

**Quer que eu:**

1. ✅ Primeiro você testa a simulação com os dados atuais (use o PASSO_A_PASSO_SIMULACAO.md)
2. ✅ Depois eu atualizo o script para capturar os vínculos facultativos
3. ✅ Você roda novamente e comparamos

**OU**

1. ❌ Já atualizo o script agora e você testa tudo de uma vez

**Qual prefere?**
