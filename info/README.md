# Sist_Prev

Sistema completo de planejamento previdenciário para RGPS (Regime Geral de Previdência Social) desenvolvido em Excel/VBA com automação Python.

**Retirement planning tool for Brazilian social security (RGPS) with comprehensive calculations, CNIS import, and benefit simulations.**

---

## 📋 Funcionalidades

### ✅ Gestão de Clientes

- Cadastro completo de dados pessoais, documentos e contatos
- Indicadores automáticos de características dos vínculos
- Sistema de busca e consulta

### ✅ Gestão de Vínculos Trabalhistas

- Cadastro manual ou importação automática via CNIS
- Detecção automática de concomitância (períodos sobrepostos)
- Suporte a vínculos especiais, rurais, militares e exterior
- Conversão automática de tempo especial (fatores 1.4, 1.75, 2.33)

### ✅ Motor de Cálculo Completo

5 regras de aposentadoria implementadas:

1. **Tempo de Contribuição** (Direito Adquirido)
2. **Aposentadoria por Idade**
3. **Regra de Pontos**
4. **Pedágio de 50%**
5. **Pedágio de 100%**

### ✅ Simulações Inteligentes

- Análise automática da melhor regra para cada cliente
- Cálculo de tempo faltante e data provável
- Estimativa do valor do benefício
- Projeções de idade na aposentadoria

### ✅ Importação de CNIS

Script Python que extrai automaticamente:

- Dados do cabeçalho (NIT, CPF, Nome, Nascimento)
- Vínculos estruturados com todas as datas
- **🆕 Remunerações por competência (mês/ano e valor)**
- Tratamento de duplicidades e vínculos em branco

**Novidade:** Com as remunerações importadas, o sistema pode calcular a média real dos 80% maiores salários para estimativa precisa do benefício.

---

## 🚀 Como Usar

### 1. Cadastrar Cliente

- Abra o formulário `frmCadastro`
- Preencha dados pessoais e documentos
- Salve

### 2. Importar Vínculos do CNIS

```bash
# Via terminal (requer Python 3.8+ e pdfplumber)
python converter_extrato_inss.py extrato_cnis.pdf saida_cnis.csv
```

**Arquivos gerados:**

- `saida_cnis_vinculos_estruturado.csv` - Vínculos (use este)
- `saida_cnis_remuneracoes.csv` - 🆕 Remunerações mensais
- `saida_cnis_dados_cliente.csv` - Dados do cabeçalho

Depois importe os arquivos CSV gerados no VBA:

```vb
' Importar vínculos
Call ImportarVinculosDeCSV("caminho/saida_cnis_vinculos_estruturado.csv", ID_Cliente)

' Importar remunerações (NOVO)
Call ImportarRemuneracoesDeCSV("caminho/saida_cnis_remuneracoes.csv", ID_Cliente)
```

### 3. Fazer Simulação

- Abra `frmSimulacoes` para o cliente
- Clique em **"Analisar Todas as Regras"**
- O sistema mostrará automaticamente a melhor opção

---

## 📊 Estrutura do Sistema

### Módulos VBA

- `modSimulacoes.bas` - Motor de cálculo (5 regras implementadas)
- `modVinculos.bas` - Gestão de vínculos e detecção de concomitância
- `modCadastro.bas` - Operações de cadastro de clientes
- `modImportacao.bas` - Importação de dados do CNIS
- `modDB.bas` - Acesso a parâmetros configuráveis
- `modUtil.bas` - Funções utilitárias

### Formulários

- `frmPrincipal.bas` - Tela principal do sistema
- `frmCadastro.bas` - Cadastro de clientes e vínculos
- `frmSimulacoes.bas` - Simulações de aposentadoria
- `frmVinculos.bas` - Edição detalhada de vínculos
- `frmBusca.bas` - Busca e consulta de clientes

### Planilhas

- `Cadastro_Clientes` - Banco de dados de clientes
- `Vinculos` - Banco de dados de vínculos trabalhistas
- `Simulacoes` - Histórico de simulações
- `Documentos` - Controle de documentos anexos
- `Config_Regras` - **Parâmetros configuráveis do sistema**

### Script Python

- `converter_extrato_inss.py` - Extração automática de dados do PDF do CNIS

---

## ⚙️ Parâmetros Configuráveis

Todos os parâmetros estão na planilha **Config_Regras** e podem ser atualizados:

| Parâmetro                | Valor Padrão | Descrição                       |
| ------------------------ | ------------ | ------------------------------- |
| Idade_Minima_Homem       | 65           | Idade mínima (homem)            |
| Idade_Minima_Mulher      | 62           | Idade mínima (mulher)           |
| Carencia_Minima          | 180          | Carência em meses               |
| Pontos_Homem             | 105          | Pontos necessários (homem)      |
| Pontos_Mulher            | 100          | Pontos necessários (mulher)     |
| Data_Reforma             | 13/11/2019   | Data da reforma previdenciária  |
| Coeficiente_Inicial      | 60           | Percentual inicial do benefício |
| Percentual_Acrescimo_Ano | 2            | Acréscimo por ano adicional     |
| Salario_Minimo           | 1412         | Salário mínimo vigente          |
| Teto_INSS                | 7786.02      | Teto do INSS                    |

---

## 🎯 Funções do Motor de Cálculo

### Cálculos Base

```vb
CalcularTempo(ID_Cliente) As Double
' Calcula tempo total de contribuição com tratamento de sobreposição

CalcularIdade(nascimento) As Long
' Calcula idade precisa considerando data de nascimento

CalcularTempoEspecial(ID_Cliente) As Double
' Converte tempo especial em tempo comum com fatores de conversão

DetectarConcomitancia(ID_Cliente)
' Identifica automaticamente vínculos sobrepostos
```

### Regras de Aposentadoria

```vb
RegraTempoContribuicao(ID_Cliente) As Collection
' 35 anos (H) ou 30 anos (M) - Direito Adquirido

RegraIdade(ID_Cliente) As Collection
' 65 anos (H) ou 62 anos (M) + carência

RegraPontos(ID_Cliente) As Collection
' 105 pontos (H) ou 100 pontos (M) + tempo mínimo

RegraPedagio50(ID_Cliente) As Collection
' Pedágio 50% para quem estava próximo da aposentadoria

RegraPedagio100(ID_Cliente) As Collection
' Pedágio 100% sem idade mínima
```

### Análise Inteligente

```vb
AnalisarMelhorRegra(ID_Cliente) As Collection
' Compara todas as regras e retorna a melhor opção

CalcularValorBeneficio(ID_Cliente, tempo) As Double
' Estima o valor do benefício com base no tempo

VerificarElegibilidadeTransicao(ID_Cliente) As String
' Verifica quais regras de transição o cliente pode usar
```

---

## 📖 Documentação Adicional

- **[MOTOR_CALCULO_STATUS.md](MOTOR_CALCULO_STATUS.md)** - Status detalhado do motor de cálculo
- **[GUIA_RAPIDO.md](GUIA_RAPIDO.md)** - Guia rápido de uso do sistema

---

## 🔧 Requisitos

### Excel/VBA

- Microsoft Excel 2016 ou superior
- Macros habilitadas

### Python (para importação de CNIS)

- Python 3.8+
- Biblioteca: `pdfplumber`

Instalação:

```bash
pip install pdfplumber
```

---

## ⚠️ Observações Importantes

### Simplificações

1. **Valor do benefício**: Usa estimativa conservadora. Para cálculo preciso é necessário:

   - Histórico completo de salários desde 1994
   - Aplicação de correção monetária
   - Cálculo da média dos 80% maiores salários

2. **Carência**: Simplificada para anos de contribuição. O ideal seria:
   - Contar número exato de contribuições mensais
   - Considerar competências pagas

### Recomendações

- ✅ Atualizar parâmetros regularmente conforme legislação
- ✅ Fazer backup semanal do arquivo
- ✅ Validar cálculos com consultoria jurídica especializada
- ✅ Testar com casos reais antes de uso em produção

---

## 📝 Status do Projeto

**Versão:** 1.0  
**Data:** 11/01/2026  
**Status:** ✅ **MOTOR DE CÁLCULO COMPLETO E OPERACIONAL**

### Implementado

- ✅ Cadastro de clientes e vínculos
- ✅ Importação automática via CNIS
- ✅ 5 regras de aposentadoria
- ✅ Cálculo de valor do benefício
- ✅ Análise automática da melhor regra
- ✅ Detecção de concomitância
- ✅ Conversão de tempo especial
- ✅ Parâmetros configuráveis

### Próximos Passos (Opcionais)

- Relatórios em PDF
- Sistema de documentos anexos
- Histórico de simulações
- Refinamento visual

---

## 📄 Licença

Sistema desenvolvido para planejamento previdenciário profissional.

---

## 🆘 Suporte

Para dúvidas ou problemas:

1. Consulte o [GUIA_RAPIDO.md](GUIA_RAPIDO.md)
2. Verifique o [MOTOR_CALCULO_STATUS.md](MOTOR_CALCULO_STATUS.md)
3. Revise os parâmetros em Config_Regras

---

**Desenvolvido com Excel VBA + Python**  
**Sistema completo de planejamento previdenciário brasileiro**
