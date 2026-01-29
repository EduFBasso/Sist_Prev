# 📊 ANÁLISE DA PLANILHA inss_simulacao.xlsx

**Data:** 26/01/2026  
**Análise de:** Planilhas/inss_simulacao.xlsx (estrutura atual do advogado)

---

## 🎯 DECISÃO: **USAR A PLANILHA EXISTENTE COM ADAPTAÇÕES**

### Razões:
1. ✅ **Estrutura sólida** - 6 abas bem organizadas + Config_Regras
2. ✅ **Lógica familiar** - Advogado já conhece o fluxo
3. ✅ **Menos trabalho** - Já tem fórmulas básicas
4. ✅ **Validação mais fácil** - Comparar com sistema VBA atual

---

## 📋 ESTRUTURA ATUAL DA PLANILHA

### Aba 1: **Config_Regras**
```
✅ APROVADA - Usar como está
- 23 parâmetros configuráveis
- Idade mínima (65H/62M)
- Carência (180 meses)
- Pontos (105H/100M)
- Fatores conversão especial (1.4, 1.75, 2.33)
- Salário mínimo (1.412)
- Teto INSS (7.786,02)
- Data reforma (13/11/2019)
- Índice correção (INPC)
```

### Aba 2: **1.Dados Gerais do Contribuinte**
```
✅ ESTRUTURA BOA - Precisa preenchimento automático

Campos atuais:
- Nome: João Carlos Eduardo Figueiredo Basso
- Data nascimento: 22/03/1973
- Sexo: Masculino
- CPF: 170.140.798-16
- Tipo aposentadoria: Comum/Especial/Mista
- Data início contribuições: (vazio - calcular do CSV)
- Data final considerada: "Hoje ou data desejada"
- Tempo total: (automático - calcular)
- Tempo especial convertido: (automático)
- Média salarial: (automático)
- Coeficiente: (automático)
- Valor final: (automático)

🔧 AÇÕES NECESSÁRIAS:
1. Preencher automaticamente do CSV dados_cliente
2. Calcular data início (MIN das competências)
3. Calcular tempo total dos vínculos
4. Implementar fórmulas de cálculo
```

### Aba 3: **2.Histórico de contribuições**
```
⚠️ PRECISA ADAPTAÇÃO - Estrutura diferente do nosso CSV

Estrutura atual (colunas):
1. Inicio (data início vínculo)
2. Fim (data fim vínculo)
3. Tipo de atividade (Comum/Especial 25/20/15)
4. Fator de conversão (1.00/1.40/1.75/2.33)
5. Tempo no período (dias) - FÓRMULA
6. Tempo convertido (dias) - FÓRMULA
7. Salário de contribuição (valor mensal)
8. Índice de correção (SELIC)
9. Salário corrigido - FÓRMULA

⚠️ PROBLEMA: Nosso CSV tem REMUNERAÇÕES por COMPETÊNCIA (MM/AAAA), não por VÍNCULO.

💡 SOLUÇÃO:
Opção A (RECOMENDADA): Manter estrutura CSV (competência)
  - Linhas = uma por competência (178 no João Carlos)
  - Colunas: Competência | Seq | CNPJ | Remuneração | Indicadores | Fator | Índice INPC | Valor Corrigido
  
Opção B: Agregar por vínculo (mais complexo)
  - Agrupar remunerações por Seq
  - Calcular média por vínculo
  - Perder granularidade mensal
```

### Aba 4: **3.Índice SELIC acumulado**
```
✅ PODE SER SUBSTITUÍDA POR INPC

Estrutura atual:
- Mês/Ano
- Índice SELIC acumulado
- Fonte
- Última atualização

💡 ALTERNATIVA:
- Usar INPC (já temos em fat_inpc_selic/)
- Script atualizar_inpc.py já gera CSV
- Importar índices INPC para esta aba
```

### Aba 5: **4.Cálculo da média salarial**
```
✅ BOA - Precisa alimentação

Estrutura:
- Coluna A: Mês/Ano
- Coluna B: Salário corrigido
- Resumo:
  * Quantidade Meses
  * Somatória Salários corrigidos
  * Média

Exemplo com 2 meses:
07/1994 | 1.235,56
08/1994 | 1.250,00
-----------------
Total: 2.485,56
Média: 1.242,78

🔧 AÇÃO:
- Popular com 178 remunerações do João Carlos
- Aplicar índice INPC em cada linha
- Fórmulas: =SOMA(B:B) e =MÉDIA(B:B)
```

### Aba 6: **5.Cálculo do coeficiente**
```
✅ BOA - Precisa fórmulas

Estrutura:
- Anos totais de contribuição
- Anos acima do mínimo (20H/15M)
- Coeficiente = 60% + (2% × anos acima)

Exemplo:
- Total: 25 anos
- Mínimo homem: 20 anos
- Acima: 5 anos
- Coeficiente: 60% + (2% × 5) = 70%

🔧 FÓRMULAS:
- B3: =SOMA(dias_convertidos)/365.25  (da aba 2)
- B4: =B3 - SE(Sexo="Masculino", 20, 15)
- B5: =0.6 + (B4 * 0.02)
```

### Aba 7: **6.Resultado**
```
✅ BOA - Precisa fórmulas finais

Estrutura:
- Média salarial corrigida (da aba 4)
- Coeficiente (da aba 5)
- Aposentadoria final = Média × Coeficiente

🔧 FÓRMULAS:
- B4: ='4.Cálculo da média salarial'!B7
- B5: ='5.Cálculo do coeficiente'!B5
- B6: =B4 * B5
```

---

## 🎯 PLANO DE AÇÃO RECOMENDADO

### FASE 1: Preparação dos Dados (Python)
```python
# Script: preparar_dados_planilha.py

1. Ler CSV remuneracoes (178 linhas João Carlos)
2. Ler índices INPC de fat_inpc_selic/
3. Para cada remuneração:
   - Competência (MM/AAAA)
   - Seq
   - CNPJ ou "FACULTATIVO"
   - Remuneração original
   - Índice INPC do mês
   - Remuneração corrigida = Original × Índice
4. Exportar para Excel (aba 2 e 4)
```

### FASE 2: Fórmulas Excel (Manual ou Python)
```excel
Aba 2 (Histórico):
  Tempo período (dias): =FIM - INICIO
  Tempo convertido: =Tempo_periodo × Fator_conversao
  Salário corrigido: =Salario × Indice_INPC

Aba 4 (Média):
  Quantidade: =CONT.VALORES(B:B)
  Somatória: =SOMA(B:B)
  Média: =B5/B3  OU  =MÉDIA(B:B)

Aba 5 (Coeficiente):
  Anos totais: =SOMA(aba2_tempo_convertido)/365.25
  Anos acima: =B3 - SE(Sexo="Masculino", 20, 15)
  Coeficiente: =0.6 + (B4 * 0.02)
  Limitador: =MÍNIMO(B5, 1.0)  [máximo 100%]

Aba 6 (Resultado):
  Média: ='4.Cálculo da média salarial'!B7
  Coeficiente: ='5.Cálculo do coeficiente'!B5
  Aposentadoria: =B4 * B5
  Limitador teto: =MÍNIMO(B6, Teto_INSS)
  Limitador piso: =MÁXIMO(B6, Salario_Minimo)
```

### FASE 3: Validação
1. Comparar resultado com sistema VBA atual
2. Verificar cada etapa:
   - ✅ Tempo total de contribuição
   - ✅ Média salarial
   - ✅ Coeficiente
   - ✅ Valor final
3. Documentar divergências (se houver)

---

## 🚀 PRÓXIMOS PASSOS (ORDENADOS)

### 1. Script Python: `preparar_dados_planilha.py`
```
Entrada:
  - saida/teste_apos_remocao_remuneracoes.csv (178 linhas)
  - fat_inpc_selic/inpc_fatores.csv (índices)
  
Processamento:
  1. Juntar remunerações + índices INPC
  2. Calcular valores corrigidos
  3. Calcular tempo de contribuição por vínculo
  
Saída:
  - Excel atualizado (abas 2, 3, 4 preenchidas)
  - OU CSV intermediários para colar manualmente
```

### 2. Implementar Fórmulas Excel
```
- Aba 1: Dados Gerais (fórmulas de sumarização)
- Aba 2: Histórico (tempo convertido, correção)
- Aba 4: Média (soma, média)
- Aba 5: Coeficiente (cálculo 60% + 2%)
- Aba 6: Resultado (fórmula final)
```

### 3. Testar com João Carlos (178 remunerações)
```
- Verificar 178 linhas importadas
- Conferir soma dos tempos
- Validar média salarial
- Comparar resultado final
```

### 4. Documentar Processo
```
- Passo-a-passo de preenchimento
- Explicação de cada fórmula
- "Por quê" de cada cálculo
- Comparação com VBA
```

---

## 💡 RECOMENDAÇÃO FINAL

**✅ USAR A PLANILHA inss_simulacao.xlsx COM ESTAS ADAPTAÇÕES:**

1. **Manter estrutura das 7 abas** (familiar para o advogado)
2. **Adaptar Aba 2** para formato de competências (não vínculos)
3. **Popular automaticamente** com Python
4. **Fórmulas Excel nativas** (transparência total)
5. **Validar contra VBA** (prova de conceito)

**Vantagens desta abordagem:**
- ✅ Advogado já conhece a estrutura
- ✅ Menos curva de aprendizado
- ✅ Comparação lado-a-lado com VBA
- ✅ Fórmulas visíveis (F2 mostra cálculo)
- ✅ Auditável célula por célula

**Próximo passo imediato:**
Criar script `preparar_dados_planilha.py` que:
1. Lê CSV de remunerações (178 linhas)
2. Aplica índices INPC
3. Gera novo Excel com dados populados
4. Pronto para inserir fórmulas manualmente

---

**Confirma esta abordagem para prosseguir?** 🎯
