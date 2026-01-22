# 🔍 ANÁLISE: Séries SELIC do Banco Central

**Data:** 22 de janeiro de 2026  
**Objetivo:** Esclarecer diferença entre séries 1178 e 4390

---

## 📊 DIFERENÇA ENTRE AS SÉRIES

### ❌ Série 1178 - INCORRETA para seu caso
**URL:** `https://api.bcb.gov.br/dados/serie/bcdata.sgs.1178/dados`

| Característica | Valor |
|----------------|-------|
| **Periodicidade** | DIÁRIA |
| **Unidade** | % acumulado no mês |
| **Valores Típicos** | 1-50% (diário acumulado) |
| **Problema** | Valores DIÁRIOS acumulam muito rápido |
| **Janeiro/1987** | 321% (acumulado diário do mês) |
| **Uso Correto** | Cálculos diários, não mensais |

**❌ ERRO:** Você estava usando esta série, por isso via valores astronômicos!

---

### ✅ Série 4390 - CORRETA (a que a advogada usa)
**URL:** `https://api.bcb.gov.br/dados/serie/bcdata.sgs.4390/dados`

| Característica | Valor |
|----------------|-------|
| **Periodicidade** | MENSAL |
| **Unidade** | % ao mês (taxa mensal) |
| **Valores Típicos** | 0.5-3% ao mês (atual), 5-20% (anos 90) |
| **Janeiro/1987** | 11.00% ao mês (razoável para época) |
| **Agosto/1986** | 2.57% ao mês |
| **Uso Correto** | ✅ Correção monetária mensal |

**✅ CORRETO:** Script `atualizar_selic_4390.py` usa esta série!

---

## 📋 COMPARAÇÃO PRÁTICA

### Exemplo: Janeiro/1987 (Hiperinflação)

| Série | Valor | Interpretação |
|-------|-------|---------------|
| **1178** | 321.15% | Acumulado DIÁRIO no mês (soma de todos os dias) |
| **4390** | 11.00% | Taxa MENSAL (% no mês inteiro) |

**Explicação:**
- 1178: Soma os juros diários → 321% acumulado
- 4390: Taxa efetiva mensal → 11% no mês

Ambos são corretos, mas **4390 é adequada para cálculos mensais**.

---

## 💰 FATORES DE CORREÇÃO - Comparação

### R$ 1.000,00 de Janeiro/1995 corrigido até Janeiro/2026:

| Índice | Fator | Valor Corrigido | Adequado? |
|--------|-------|-----------------|-----------|
| **INPC** (188) | 13.52x | R$ 13.520,00 | ✅ Previdenciário |
| **SELIC** (4390) | 75.55x | R$ 75.546,10 | ⚠️ Judiciais |
| SELIC (1178) | 49 quatrilhões | ❌ ERRO | ❌ Série errada |

---

## 🎯 RECOMENDAÇÕES FINAIS

### Para o Sistema Sist_Prev:

1. **Correção de Salários Previdenciários:**
   - ✅ Usar **INPC** (série 188)
   - Script: `atualizar_inpc.py`
   - Motivo: É o índice legal para previdência

2. **Se advogada REALMENTE exigir SELIC:**
   - ✅ Usar **série 4390** (mensal)
   - Script: `atualizar_selic_4390.py`
   - ❌ NUNCA usar série 1178 (diária)

3. **Arquivos Gerados:**
   ```
   saida/inpc_fatores.csv           ← INPC (recomendado)
   saida/selic_fatores.csv          ← SELIC 4390 (se advogada exigir)
   saida/selic_bruto_para_excel.csv ← Dados brutos (formato VBA)
   ```

---

## 📝 SCRIPT VBA ANALISADO

```vb
url = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.4390/dados?formato=json"
```

**✅ CORRETO!** A advogada está usando a série 4390 (mensal).

### Integração com seu sistema:

O arquivo `selic_bruto_para_excel.csv` tem EXATAMENTE o formato que o VBA dela espera:
```
Data;Valor;Fonte
01/08/1986;2.570000;BCB API
01/09/1986;2.940000;BCB API
...
```

Pode importar direto na planilha "3.Indice SELIC acumulado"!

---

## ⚖️ DECISÃO FINAL: INPC ou SELIC?

### Argumentos para INPC:
- ✅ Índice oficial para correção previdenciária
- ✅ Lei exige INPC para revisões
- ✅ INSS usa INPC
- ✅ Jurisprudência consolidada

### Argumentos para SELIC:
- ⚠️ Apenas para juros de mora em ações judiciais
- ⚠️ NÃO é para correção de salários
- ⚠️ Pode inflar valores incorretamente

### Sugestão:
**Converse com a advogada:**
- "SELIC é para corrigir salários do CNIS ou apenas juros de mora judicial?"
- "Se for juros de mora, faz sentido. Se for correção de salários, o correto é INPC."

---

## 📞 PRÓXIMOS PASSOS

1. ✅ Validar com advogada: INPC ou SELIC 4390?
2. ✅ Se SELIC: usar `atualizar_selic_4390.py`
3. ✅ Se INPC: usar `atualizar_inpc.py`
4. ❌ NUNCA usar série 1178!

---

**Última atualização:** 22/01/2026  
**Script correto criado:** `atualizar_selic_4390.py` ✅
