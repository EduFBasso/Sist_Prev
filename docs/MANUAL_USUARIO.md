# Manual do Usuário - Sistema de Planejamento Previdenciário RGPS

## Sumário

1. [Introdução](#introdução)
2. [Cadastro de Clientes](#cadastro-de-clientes)
3. [Importação de Vínculos](#importação-de-vínculos)
4. [Simulação de Aposentadorias](#simulação-de-aposentadorias)
5. [Interpretação dos Resultados](#interpretação-dos-resultados)
6. [Tipos de Aposentadoria](#tipos-de-aposentadoria)

## Introdução

O Sistema de Planejamento Previdenciário foi desenvolvido para auxiliar advogados e profissionais previdenciários na análise de elegibilidade para diferentes modalidades de aposentadoria do RGPS.

### Funcionalidades Principais

- ✓ Cadastro completo de clientes
- ✓ Validação automática de CPF
- ✓ Importação de vínculos do INSS (CSV)
- ✓ Cálculo automático de tempos de contribuição
- ✓ Simulação de 4 tipos de aposentadoria
- ✓ Relatórios detalhados

## Cadastro de Clientes

### Usando a Interface VBA

1. Pressione `Alt + F8` para abrir a lista de macros
2. Execute a macro `ExemploCadastrarCliente` ou crie uma personalizada
3. Modifique o código com os dados do cliente:

```vba
ModClientes.CadastrarCliente _
    nome:="Nome Completo do Cliente", _
    cpf:="000.000.000-00", _
    dataNascimento:=DateSerial(1970, 1, 15), _
    sexo:="M", _  ' M para masculino, F para feminino
    email:="cliente@email.com", _
    telefone:="(00) 00000-0000"
```

### Cadastro Manual na Planilha

1. Acesse a planilha **Clientes**
2. Preencha uma nova linha com os dados:
   - **ID**: Será gerado automaticamente
   - **Nome**: Nome completo
   - **CPF**: Formato 000.000.000-00
   - **Data Nascimento**: Formato dd/mm/aaaa
   - **Sexo**: M ou F
   - **Email**: Endereço de e-mail
   - **Telefone**: Número de contato
   - **Data Cadastro**: Será preenchida automaticamente

### Validações Automáticas

O sistema valida:
- ✓ Formato e dígitos verificadores do CPF
- ✓ CPF não duplicado
- ✓ Campos obrigatórios preenchidos

## Importação de Vínculos

### Método 1: Importação via CSV

#### Preparar o Arquivo CSV

Crie um arquivo de texto com extensão `.csv` contendo os vínculos, separados por ponto-e-vírgula:

```
Empresa;CNPJ;Data Início;Data Fim;Tipo Vínculo;Condição Especial
Empresa ABC Ltda;12.345.678/0001-90;01/01/2000;31/12/2010;CLT;Normal
Empresa XYZ SA;98.765.432/0001-10;01/01/2011;31/12/2020;CLT;Insalubre
Empresa 123 Ltda;11.222.333/0001-44;01/01/2021;;CLT;Normal
```

**Notas:**
- Data Fim vazia indica vínculo ativo
- Condição Especial pode ser: Normal, Insalubre, Periculoso, Perigoso

#### Executar a Importação

```vba
ModImportacaoINSS.ImportarVinculosCSV _
    cpfCliente:="000.000.000-00", _
    caminhoArquivo:="C:\caminho\do\arquivo.csv"
```

### Método 2: Cadastro Manual de Vínculos

```vba
ModImportacaoINSS.AdicionarVinculo _
    cpfCliente:="000.000.000-00", _
    empresa:="Nome da Empresa", _
    cnpj:="00.000.000/0001-00", _
    dataInicio:=DateSerial(2000, 1, 1), _
    dataFim:=DateSerial(2020, 12, 31), _
    tipoVinculo:="CLT", _
    condicaoEspecial:="Normal"
```

### Método 3: Preenchimento Direto na Planilha

1. Acesse a planilha **Vínculos**
2. Adicione uma nova linha com os dados do vínculo
3. Execute `CalcularTemposContribuicao` para atualizar os cálculos

## Simulação de Aposentadorias

### Executar Simulação

```vba
' Método 1: Via código VBA
Dim resultado As String
resultado = ModSimulacaoAposentadoria.SimularAposentadorias("000.000.000-00")
MsgBox resultado

' Método 2: Executar macro
' Pressione Alt + F8 e execute ExemploSimularAposentadoria
```

### Cálculos Automáticos

O sistema calcula automaticamente:
- ✓ Idade atual do cliente
- ✓ Tempo total de contribuição
- ✓ Tempo de contribuição especial
- ✓ Pontos para aposentadoria por pontos
- ✓ Requisitos faltantes para cada modalidade

## Interpretação dos Resultados

### Símbolos nos Resultados

- **✓** = Requisitos atendidos, cliente pode se aposentar
- **✗** = Requisitos não atendidos, veja o que falta

### Planilha ResultadosSimulacao

Consulte esta planilha para ver um resumo de todas as simulações:

| Campo | Descrição |
|-------|-----------|
| CPF Cliente | Identificação do cliente |
| Tipo Aposentadoria | Modalidade analisada |
| Elegível | Sim/Não |
| Anos Faltantes | Tempo que falta para atingir requisitos |
| Data Simulação | Quando foi realizada a análise |

## Tipos de Aposentadoria

### 1. Aposentadoria por Idade

**Requisitos (pós-reforma 2019):**
- Homens: 65 anos de idade + 20 anos de contribuição
- Mulheres: 62 anos de idade + 15 anos de contribuição

**Quando usar:** Cliente próximo da idade mínima mas com pouco tempo de contribuição.

### 2. Aposentadoria por Tempo de Contribuição

**Requisitos (regra de transição - pedágio 100%):**
- Homens: 35 anos de contribuição + 60 anos de idade
- Mulheres: 30 anos de contribuição + 57 anos de idade

**Quando usar:** Cliente com muito tempo de contribuição mas ainda jovem.

### 3. Aposentadoria por Pontos

**Requisitos (progressivos):**
- Homens: 35 anos de contribuição + pontos (100 em 2019, até 105 em 2028)
- Mulheres: 30 anos de contribuição + pontos (90 em 2019, até 100 em 2033)

**Cálculo dos pontos:** Idade + Tempo de contribuição

**Quando usar:** Melhor alternativa para quem tem bom equilíbrio entre idade e tempo.

### 4. Aposentadoria Especial

**Requisitos:**
- 15, 20 ou 25 anos de atividade especial (conforme grau de risco)
- Sistema considera 25 anos por padrão

**Condições especiais:**
- Insalubridade
- Periculosidade
- Penosidade
- Exposição a agentes nocivos

**Quando usar:** Cliente trabalhou exposto a condições prejudiciais à saúde.

## Dicas de Uso

### Melhores Práticas

1. **Sempre atualize os dados antes de simular**
   - Verifique se todos os vínculos estão cadastrados
   - Confirme que as datas estão corretas

2. **Simule regularmente**
   - Refaça a simulação periodicamente
   - Acompanhe a evolução dos requisitos

3. **Documente casos especiais**
   - Anote observações importantes
   - Guarde cópias dos extratos do INSS

4. **Valide com o CNIS**
   - Compare com o Cadastro Nacional de Informações Sociais
   - Identifique períodos não reconhecidos

### Atalhos Úteis

- `Alt + F8`: Lista de macros
- `Alt + F11`: Editor VBA
- `Ctrl + PageDown/PageUp`: Navegar entre planilhas

## Casos Especiais

### Conversão de Tempo Especial

Para converter tempo especial em tempo comum:
- Multiplicar por 1,4 (homens) ou 1,2 (mulheres)
- Implementação futura planejada

### Múltiplos Vínculos Simultâneos

- O sistema soma todos os períodos
- Importante: INSS pode ter regras específicas
- Valide com análise jurídica

### Períodos Rurais

- Cadastre como vínculos normais
- Use "Rural" no campo Tipo Vínculo
- Atenção às regras especiais do INSS

## Suporte e Atualizações

Para suporte técnico ou sugestões de melhorias, consulte o repositório GitHub do projeto.

---

**Versão:** 1.0  
**Última atualização:** Janeiro 2026
