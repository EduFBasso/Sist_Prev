# Guia Rápido - Sist_Prev

## 🚀 COMO USAR O SISTEMA

### 1. CADASTRAR UM CLIENTE

1. Abra o formulário **frmCadastro**
2. Preencha os dados pessoais:
   - Nome, CPF, PIS, RG
   - Data de nascimento, Sexo
   - Contatos e endereço
3. Clique em **Salvar**

### 2. IMPORTAR VÍNCULOS DO CNIS

#### Opção A: Via Python (Automático)

```bash
# No terminal, dentro da pasta do projeto:
python converter_extrato_inss.py arquivo_cnis.pdf saida_cnis.csv
```

Isso gerará 4 arquivos CSV:

- `saida_cnis.csv` - Dados brutos
- `saida_cnis_dados_cliente.csv` - Dados do cabeçalho
- `saida_cnis_vinculos_brutos.csv` - Vínculos brutos
- `saida_cnis_vinculos_estruturado.csv` - **Use este para importar**

#### Opção B: No VBA

```vb
' No módulo de importação:
Call ImportarVinculosDeCSV("caminho/saida_cnis_vinculos_estruturado.csv", ID_Cliente)
```

### 3. CADASTRAR VÍNCULOS MANUALMENTE

1. No formulário do cliente, clique em **Adicionar Vínculo**
2. Preencha:
   - Data Início e Data Fim
   - Tipo de vínculo (Empregado, Autônomo, etc.)
   - Se é especial: marque "Sim" e informe o grau (15, 20 ou 25)
   - Outros indicadores (Rural, Militar, Exterior, etc.)
3. Clique em **Salvar**

### 4. FAZER SIMULAÇÕES DE APOSENTADORIA

1. Abra o formulário **frmSimulacoes** para o cliente desejado
2. O sistema mostrará automaticamente:

   - Idade atual
   - Tempo total de contribuição
   - Tempo especial convertido
   - Pontos atuais

3. **Opção A: Análise Automática (Recomendado)**

   - Clique em **Analisar Todas as Regras**
   - O sistema testará todas as 5 regras
   - Selecionará automaticamente a melhor opção
   - Mostrará os resultados

4. **Opção B: Escolha Manual**

   - Marque a regra desejada:
     - ⚪ Tempo de Contribuição (Direito Adquirido)
     - ⚪ Idade
     - ⚪ Pontos
     - ⚪ Pedágio 50%
     - ⚪ Pedágio 100%
   - Clique em **Calcular**

5. **Resultados Exibidos:**
   - ✅ **Direito**: Se já pode se aposentar (Sim/Não)
   - ⏱️ **Falta**: Quantos anos ainda faltam
   - 📅 **Data Provável**: Quando poderá se aposentar
   - 👤 **Idade Projetada**: Idade na data da aposentadoria
   - 💰 **Valor Estimado**: Valor estimado do benefício
   - 📝 **Observações**: Requisitos e pendências

### 5. ENTENDER OS RESULTADOS

#### Exemplo 1: Cliente com Direito

```
Direito: Sim
Falta: 0.00 anos
Data Provável: 11/01/2026 (hoje)
Idade Projetada: 58 anos
Valor Estimado: R$ 5.450,00
Obs: Direito adquirido pela regra de pontos.
```

✅ Pode se aposentar AGORA!

#### Exemplo 2: Cliente sem Direito

```
Direito: Não
Falta: 2.50 anos
Data Provável: 11/07/2028
Idade Projetada: 60 anos
Valor Estimado: R$ 6.200,00
Obs: Falta pontuação.
```

❌ Ainda faltam 2 anos e 6 meses

## 🎯 AS 5 REGRAS EXPLICADAS

### 1️⃣ Tempo de Contribuição (Direito Adquirido)

**Quem pode usar:** Quem completou os requisitos ANTES de 13/11/2019

- **Homem:** 35 anos de contribuição
- **Mulher:** 30 anos de contribuição
- **Sem idade mínima**

### 2️⃣ Aposentadoria por Idade

**Regra atual pós-reforma**

- **Homem:** 65 anos + 180 meses de carência
- **Mulher:** 62 anos + 180 meses de carência

### 3️⃣ Regra de Pontos

**Pontos = Idade + Tempo de Contribuição**

- **Homem:** 105 pontos + 35 anos de contribuição
- **Mulher:** 100 pontos + 30 anos de contribuição
- Exemplo: Homem com 58 anos e 35 anos de contribuição = 93 pontos (precisa de mais 12 pontos)

### 4️⃣ Pedágio de 50%

**Quem pode usar:** Quem estava a ATÉ 2 ANOS da aposentadoria em 13/11/2019

- Deve cumprir o tempo que faltava + 50% desse tempo
- Exemplo: Faltavam 1,5 anos → deve trabalhar 1,5 + 0,75 = 2,25 anos
- **Sem idade mínima**

### 5️⃣ Pedágio de 100%

**Quem pode usar:** Qualquer pessoa que já contribuía antes de 13/11/2019

- Deve cumprir o tempo que faltava + 100% desse tempo
- Exemplo: Faltavam 3 anos → deve trabalhar 3 + 3 = 6 anos
- **Sem idade mínima**

## 📊 PARÂMETROS DO SISTEMA

Os parâmetros estão na planilha **Config_Regras** e podem ser atualizados:

| Parâmetro            | Como Atualizar                                  |
| -------------------- | ----------------------------------------------- |
| Salário Mínimo       | Atualizar quando houver reajuste oficial        |
| Teto INSS            | Atualizar quando houver reajuste oficial        |
| Idade Mínima         | Verificar progressão anual da idade (se houver) |
| Pontos               | Verificar progressão anual (se houver)          |
| Fatores de Conversão | Normalmente fixos, mas verificar legislação     |

## ⚠️ OBSERVAÇÕES IMPORTANTES

### Vínculos Especiais

- **Grau 25 anos:** Agentes químicos, físicos ou biológicos
- **Grau 20 anos:** Atividades de risco moderado
- **Grau 15 anos:** Atividades de risco alto (mineração, etc.)

O sistema converte automaticamente tempo especial em comum usando os fatores:

- 25 anos → multiplica por 1.4
- 20 anos → multiplica por 1.75
- 15 anos → multiplica por 2.33

### Vínculos Concomitantes

O sistema detecta automaticamente períodos sobrepostos e considera apenas uma vez no cálculo do tempo total.

### Valor do Benefício

O valor exibido é uma **ESTIMATIVA**. Para cálculo preciso:

- É necessário o histórico completo de salários desde 1994
- Aplicação de correção monetária
- Cálculo da média dos 80% maiores salários

## 🔧 MANUTENÇÃO

### Atualizar Parâmetros Anuais

1. Abrir planilha **Config_Regras**
2. Atualizar valores na coluna "Valor"
3. Atualizar data na coluna "Ultima_Atualizacao"
4. Salvar

### Backup Recomendado

- Fazer backup semanal do arquivo .xlsm
- Manter histórico de simulações importantes
- Documentar mudanças na legislação

## 🆘 TROUBLESHOOTING

### Erro: "Cliente não encontrado"

- Verifique se o cliente foi salvo corretamente
- Confirme que o ID está correto

### Erro: "Selecione uma regra"

- Marque uma das opções de regra de aposentadoria
- Ou use o botão "Analisar Todas"

### Vínculos não aparecem

- Verifique se o ID_Cliente está correto
- Confirme que os vínculos foram salvos
- Verifique se há datas inválidas

### Tempo de contribuição zerado

- Verifique se há vínculos cadastrados
- Confirme que as datas de início e fim são válidas
- Verifique formato das datas (dd/mm/aaaa)

---

**Versão do Sistema:** 1.0  
**Data:** 11/01/2026  
**Status:** Sistema Completo e Operacional
