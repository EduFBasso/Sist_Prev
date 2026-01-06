# Sistema de Planejamento Previdenciário - RGPS

Sistema completo de planejamento previdenciário voltado ao RGPS (Regime Geral de Previdência Social), desenvolvido em Excel/VBA para uso em escritórios de advocacia previdenciária.

## 📋 Descrição

Este sistema permite cadastrar clientes, importar vínculos empregatícios a partir de extratos do INSS, calcular tempos de contribuição e simular diferentes regras de aposentadoria, apresentando resultados de forma acessível e profissional.

## ✨ Funcionalidades

### Gerenciamento de Clientes
- ✅ Cadastro completo de clientes com validação de CPF
- ✅ Armazenamento de dados pessoais (nome, CPF, data de nascimento, sexo, contatos)
- ✅ Busca e listagem de clientes
- ✅ Validação automática de duplicidade

### Importação de Vínculos
- ✅ Importação de vínculos via arquivo CSV
- ✅ Cadastro manual de vínculos empregatícios
- ✅ Registro de condições especiais (insalubridade, periculosidade)
- ✅ Suporte para vínculos ativos e encerrados

### Cálculo de Tempos
- ✅ Cálculo automático de tempo total de contribuição
- ✅ Cálculo de tempo especial separado
- ✅ Conversão automática de dias para anos/meses/dias
- ✅ Histórico de cálculos com data/hora

### Simulação de Aposentadorias
- ✅ **Aposentadoria por Idade** (pós-reforma 2019)
- ✅ **Aposentadoria por Tempo de Contribuição** (regra de transição)
- ✅ **Aposentadoria por Pontos** (progressiva)
- ✅ **Aposentadoria Especial**
- ✅ Indicação clara de requisitos atendidos ou faltantes
- ✅ Cálculo automático de tempo/idade faltante

### Relatórios e Resultados
- ✅ Planilha de resultados com todas as simulações
- ✅ Dashboard principal com instruções
- ✅ Interface amigável e profissional
- ✅ Exportação facilitada para relatórios

## 🚀 Instalação

### Requisitos
- Microsoft Excel 2013 ou superior
- Windows 7 ou superior / macOS com Excel
- Macros habilitadas no Excel

### Passo a Passo

1. **Baixe o repositório**
   ```bash
   git clone https://github.com/EduFBasso/Sist_Prev.git
   ```

2. **Crie uma pasta de trabalho Excel**
   - Abra o Microsoft Excel
   - Crie nova pasta de trabalho
   - Salve como `Sistema_Previdenciario.xlsm` (formato com macros)

3. **Importe os módulos VBA**
   - Pressione `Alt + F11` para abrir o Editor VBA
   - Importe os arquivos da pasta `src/modules/`:
     - `ModPrincipal.bas`
     - `ModClientes.bas`
     - `ModImportacaoINSS.bas`
     - `ModSimulacaoAposentadoria.bas`

4. **Inicialize o sistema**
   - Pressione `Alt + F8`
   - Execute a macro `InicializarSistema`
   - As planilhas serão criadas automaticamente

📖 Consulte o [Manual de Instalação Completo](docs/INSTALACAO.md) para instruções detalhadas.

## 📚 Documentação

- **[Manual de Instalação](docs/INSTALACAO.md)** - Guia completo de instalação
- **[Manual do Usuário](docs/MANUAL_USUARIO.md)** - Instruções detalhadas de uso
- **[Exemplo de CSV](templates/exemplo_vinculos.csv)** - Template para importação de vínculos

## 💡 Uso Rápido

### Cadastrar Cliente
```vba
ModClientes.CadastrarCliente _
    nome:="João da Silva", _
    cpf:="123.456.789-00", _
    dataNascimento:=DateSerial(1965, 5, 15), _
    sexo:="M", _
    email:="joao@email.com", _
    telefone:="(11) 98765-4321"
```

### Adicionar Vínculo
```vba
ModImportacaoINSS.AdicionarVinculo _
    cpfCliente:="123.456.789-00", _
    empresa:="Empresa ABC Ltda", _
    cnpj:="12.345.678/0001-90", _
    dataInicio:=DateSerial(2000, 1, 1), _
    dataFim:=DateSerial(2020, 12, 31), _
    tipoVinculo:="CLT", _
    condicaoEspecial:="Normal"
```

### Simular Aposentadorias
```vba
Dim resultado As String
resultado = ModSimulacaoAposentadoria.SimularAposentadorias("123.456.789-00")
MsgBox resultado
```

### Demonstração Completa
Execute a macro `DemonstracaoCompleta` para ver o sistema em ação com dados de exemplo.

## 📊 Estrutura do Sistema

### Planilhas Criadas
- **Início** - Dashboard principal com instruções
- **Clientes** - Cadastro de clientes
- **Vínculos** - Registro de vínculos empregatícios
- **CalculoTempo** - Cálculos de tempo de contribuição
- **ResultadosSimulacao** - Resultados das simulações

### Módulos VBA
- **ModPrincipal** - Inicialização e coordenação
- **ModClientes** - Gestão de clientes e validações
- **ModImportacaoINSS** - Importação e gestão de vínculos
- **ModSimulacaoAposentadoria** - Simulações de aposentadoria

## 🎯 Tipos de Aposentadoria Suportados

### 1. Aposentadoria por Idade
- **Homens:** 65 anos + 20 anos de contribuição
- **Mulheres:** 62 anos + 15 anos de contribuição

### 2. Aposentadoria por Tempo de Contribuição
- **Homens:** 35 anos de contribuição + 60 anos de idade
- **Mulheres:** 30 anos de contribuição + 57 anos de idade

### 3. Aposentadoria por Pontos (Progressiva)
- **Homens:** 35 anos + pontos (100-105)
- **Mulheres:** 30 anos + pontos (90-100)
- **Pontos** = Idade + Tempo de Contribuição

### 4. Aposentadoria Especial
- 15, 20 ou 25 anos de atividade especial
- Condições: Insalubridade, Periculosidade, Penosidade

## 📄 Formato de Importação CSV

O arquivo CSV deve ter as seguintes colunas separadas por ponto-e-vírgula (;):

```csv
Empresa;CNPJ;Data Início;Data Fim;Tipo Vínculo;Condição Especial
Empresa ABC Ltda;12.345.678/0001-90;01/01/2010;31/12/2015;CLT;Normal
Empresa XYZ SA;98.765.432/0001-10;01/01/2016;;CLT;Insalubre
```

**Nota:** Data Fim vazia indica vínculo ativo.

## 🔒 Segurança

- Validação rigorosa de CPF com dígitos verificadores
- Verificação de duplicidade de cadastros
- Tratamento de erros em todas as operações
- Backup recomendado antes de importações em lote

## 🛠️ Tecnologias

- **Microsoft Excel** - Plataforma base
- **VBA (Visual Basic for Applications)** - Linguagem de programação
- **CSV** - Formato de importação de dados

## 📝 Changelog

### Versão 1.0 (Janeiro 2026)
- ✅ Implementação inicial do sistema
- ✅ Cadastro de clientes com validação de CPF
- ✅ Importação de vínculos via CSV
- ✅ Cálculo automático de tempos de contribuição
- ✅ Simulação de 4 tipos de aposentadoria
- ✅ Dashboard e documentação completa

## 🤝 Contribuindo

Contribuições são bem-vindas! Por favor:

1. Faça um fork do projeto
2. Crie uma branch para sua feature (`git checkout -b feature/MinhaFeature`)
3. Commit suas mudanças (`git commit -m 'Adiciona MinhaFeature'`)
4. Push para a branch (`git push origin feature/MinhaFeature`)
5. Abra um Pull Request

## 📧 Suporte

Para questões, sugestões ou problemas:
- Abra uma [Issue](https://github.com/EduFBasso/Sist_Prev/issues)
- Consulte a [Documentação](docs/)

## 📜 Licença

Este projeto está sob a licença MIT. Veja o arquivo LICENSE para mais detalhes.

## ⚖️ Nota Legal

Este sistema é uma ferramenta de apoio e não substitui a análise jurídica profissional. As simulações são baseadas nas regras vigentes da legislação previdenciária brasileira, mas casos específicos podem ter particularidades que requerem avaliação individualizada.

---

**Desenvolvido para planejamento previdenciário RGPS**  
**Versão:** 1.0  
**Última atualização:** Janeiro 2026
