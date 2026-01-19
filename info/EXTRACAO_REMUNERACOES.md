# Extração de Remunerações do CNIS

## 📊 Nova Funcionalidade Implementada

O script Python agora extrai também as **tabelas de remunerações** de cada vínculo do CNIS, permitindo cálculos mais precisos da média salarial.

## 🔍 O que é extraído

Para cada vínculo (empresa), o sistema extrai a tabela de remunerações com:

- **Competência**: Mês/Ano (ex: 01/1995, 02/1995)
- **Remuneração**: Valor recebido naquela competência
- **Indicadores**: Informações adicionais (se houver)

### Exemplo de Dados Extraídos

**Vínculo:**

- Seq: 1
- Código Emp: 56.528.946/0001-80
- Empresa: EMBIARA SERVICOS EMPRESARIAIS LTDA

**Remunerações:**

```
Competência | Remuneração | Indicadores
01/1995     | 286,25      |
02/1995     | 286,25      |
03/1995     | 286,25      |
```

## 🚀 Como Usar

### 1. Executar o Script Python

```bash
python converter_extrato_inss.py extrato_cnis.pdf saida_cnis.csv
```

**Arquivos gerados:**

- `saida_cnis.csv` - Dados brutos
- `saida_cnis_dados_cliente.csv` - Dados do cabeçalho
- `saida_cnis_vinculos_brutos.csv` - Vínculos brutos
- `saida_cnis_vinculos_estruturado.csv` - Vínculos estruturados
- **`saida_cnis_remuneracoes.csv` - 🆕 Remunerações por vínculo**

### 2. Importar no Excel/VBA

#### a) Importar Vínculos (como antes)

```vb
Call ImportarVinculosDeCSV("caminho/saida_cnis_vinculos_estruturado.csv", ID_Cliente)
```

#### b) Importar Remunerações (NOVO)

```vb
Call ImportarRemuneracoesDeCSV("caminho/saida_cnis_remuneracoes.csv", ID_Cliente)
```

### 3. Calcular Média Salarial

Com as remunerações importadas, agora é possível calcular a média dos 80% maiores salários:

```vb
Dim mediaSalarial As Double
mediaSalarial = CalcularMediaSalarios(ID_Cliente)
MsgBox "Média salarial: R$ " & Format(mediaSalarial, "#,##0.00")
```

## 📋 Estrutura dos Dados

### CSV de Remunerações

**Colunas:**

- `Pagina` - Número da página do PDF
- `Seq` - Sequência do vínculo
- `CodigoEmp` - CNPJ da empresa
- `Competencia` - Mês/Ano (mm/aaaa)
- `Remuneracao` - Valor numérico
- `Indicadores` - Texto adicional (se houver)

**Exemplo:**

```csv
Pagina;Seq;CodigoEmp;Competencia;Remuneracao;Indicadores
1;1;56.528.946/0001-80;01/1995;286.25;
1;1;56.528.946/0001-80;02/1995;286.25;
1;1;56.528.946/0001-80;03/1995;286.25;
```

### Planilha Remuneracoes (Excel)

**Colunas:**

- `ID_Remuneracao` - ID único
- `ID_Vinculo` - Referência ao vínculo
- `ID_Cliente` - Referência ao cliente
- `Competencia` - Mês/Ano
- `Valor` - Valor numérico
- `Indicadores` - Informações adicionais
- `Seq` - Sequência (para referência)
- `CodigoEmp` - CNPJ (para referência)

## 🎯 Benefícios

### 1. Cálculo Preciso do Benefício

Antes: Estimativa baseada em 60% do teto  
Agora: Cálculo real da média dos 80% maiores salários

### 2. Histórico Completo

- Visualização de toda trajetória salarial
- Identificação de períodos com salários baixos
- Análise de evolução salarial

### 3. Validação de Dados

- Comparação entre última remuneração e data fim do vínculo
- Identificação de inconsistências
- Detecção de períodos sem remuneração

## 🔧 Funções VBA Implementadas

### ImportarRemuneracoesDeCSV()

Importa as remunerações do CSV para a planilha Remuneracoes.

**Funcionalidades:**

- Cria automaticamente a planilha se não existir
- Associa cada remuneração ao vínculo correto
- Evita duplicatas
- Exibe total de registros importados

### CalcularMediaSalarios()

Calcula a média dos 80% maiores salários conforme regra do INSS.

**Funcionalidades:**

- Coleta todos os salários do cliente
- Ordena em ordem decrescente
- Seleciona os 80% maiores
- Retorna a média

## 📊 Exemplo de Uso Completo

```vb
Sub ProcessarClienteCNIS()
    Dim ID_Cliente As Long
    Dim pasta As String
    Dim mediaSalarial As Double
    Dim tempoTotal As Double
    Dim valorBeneficio As Double

    ID_Cliente = 1
    pasta = "C:\Dados\CNIS\"

    ' 1. Importar vínculos
    Call ImportarVinculosDeCSV(pasta & "saida_cnis_vinculos_estruturado.csv", ID_Cliente)

    ' 2. Importar remunerações
    Call ImportarRemuneracoesDeCSV(pasta & "saida_cnis_remuneracoes.csv", ID_Cliente)

    ' 3. Calcular média salarial real
    mediaSalarial = CalcularMediaSalarios(ID_Cliente)
    MsgBox "Média salarial: R$ " & Format(mediaSalarial, "#,##0.00")

    ' 4. Calcular tempo de contribuição
    tempoTotal = CalcularTempo(ID_Cliente)

    ' 5. Calcular valor do benefício com média real
    ' (Aqui você pode criar uma versão melhorada do CalcularValorBeneficio
    ' que aceita a média salarial como parâmetro)

    MsgBox "Processamento completo!"
End Sub
```

## ⚠️ Observações Importantes

### Associação Vínculo-Remuneração

O sistema associa remunerações aos vínculos através do **Código da Empresa (CNPJ)** que é salvo nas observações do vínculo durante a importação.

### Valores Monetários

- O CSV usa ponto (.) como separador decimal
- No Excel, os valores são convertidos para tipo numérico
- Use `Format()` para exibição formatada

### Performance

Para clientes com muitos vínculos e remunerações:

- A importação pode levar alguns segundos
- A planilha Remuneracoes pode ficar grande
- Considere criar índices ou filtros

### Correção Monetária

⚠️ **IMPORTANTE:** Os valores das remunerações NÃO estão corrigidos monetariamente.

Para cálculo oficial do benefício, é necessário:

1. Aplicar correção monetária usando índices oficiais
2. Converter valores de planos econômicos antigos
3. Atualizar para valores de hoje

## 🔄 Próximos Passos (Opcional)

### Melhorias Possíveis

1. **Correção Monetária Automática**

   - Tabela de índices de correção
   - Função que atualiza valores históricos

2. **Gráficos de Evolução**

   - Gráfico de linha mostrando evolução salarial
   - Identificação de períodos críticos

3. **Análise de Gaps**

   - Identificar períodos sem remuneração
   - Sinalizar inconsistências com datas de vínculo

4. **Relatório de Remunerações**
   - PDF com histórico completo
   - Análise estatística (média, mediana, máximo, mínimo)

## 📖 Referências

### Regra dos 80% Maiores Salários (INSS)

De acordo com a legislação previdenciária:

- Considera-se a média aritmética simples dos maiores salários de contribuição
- Correspondentes a 80% de todo o período contributivo desde julho de 1994
- Valores devem ser corrigidos monetariamente até a data do cálculo

---

**Implementado em:** 11/01/2026  
**Status:** ✅ Funcional - Extração e importação completas  
**Pendente:** Correção monetária automática (opcional)
