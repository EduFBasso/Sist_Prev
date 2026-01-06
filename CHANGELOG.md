# Changelog

Todas as mudanças notáveis neste projeto serão documentadas neste arquivo.

O formato é baseado em [Keep a Changelog](https://keepachangelog.com/pt-BR/1.0.0/),
e este projeto adere ao [Semantic Versioning](https://semver.org/lang/pt-BR/).

## [1.0.0] - 2026-01-06

### Adicionado

#### Sistema Core
- Sistema completo de planejamento previdenciário RGPS em Excel/VBA
- Estrutura modular com 4 módulos VBA principais
- Sistema de inicialização automática de planilhas

#### Gerenciamento de Clientes
- Cadastro completo de clientes (nome, CPF, data nascimento, sexo, contatos)
- Validação rigorosa de CPF com dígitos verificadores
- Verificação de duplicidade de cadastros
- Busca de clientes por CPF
- Listagem de todos os clientes cadastrados

#### Importação de Vínculos
- Importação de vínculos via arquivo CSV
- Cadastro manual de vínculos empregatícios
- Suporte para diferentes tipos de vínculo (CLT, Estatutário, Autônomo, MEI, Rural, Doméstico)
- Registro de condições especiais (Normal, Insalubre, Periculoso, Perigoso, Penoso)
- Suporte para vínculos ativos (sem data fim)

#### Cálculo de Tempos
- Cálculo automático de tempo total de contribuição
- Cálculo separado de tempo especial
- Conversão automática de dias para anos/meses/dias
- Histórico de cálculos com timestamp

#### Simulação de Aposentadorias
- **Aposentadoria por Idade:** Regras pós-reforma 2019
  - Homens: 65 anos + 20 anos contribuição
  - Mulheres: 62 anos + 15 anos contribuição
- **Aposentadoria por Tempo de Contribuição:** Regra de transição (pedágio 100%)
  - Homens: 35 anos contribuição + 60 anos idade
  - Mulheres: 30 anos contribuição + 57 anos idade
- **Aposentadoria por Pontos:** Sistema progressivo
  - Homens: 35 anos + pontos (100-105)
  - Mulheres: 30 anos + pontos (90-100)
  - Pontuação aumenta 1 ponto por ano
- **Aposentadoria Especial:** 15, 20 ou 25 anos conforme grau
- Indicação clara de requisitos atendidos (✓) ou não atendidos (✗)
- Cálculo automático de tempo/idade faltante
- Armazenamento de resultados em planilha dedicada

#### Interface e Planilhas
- **Planilha Início:** Dashboard principal com instruções e menu rápido
- **Planilha Clientes:** Cadastro formatado com validações
- **Planilha Vínculos:** Registro de vínculos empregatícios
- **Planilha CalculoTempo:** Resultados dos cálculos automáticos
- **Planilha ResultadosSimulacao:** Histórico de todas as simulações
- Cabeçalhos formatados profissionalmente
- Cores institucionais consistentes (RGB 68, 114, 196)

#### Documentação
- README.md completo em português com instruções detalhadas
- Manual de Instalação passo a passo
- Manual do Usuário com exemplos práticos
- Guia de Importação de Extratos do INSS
- Guia de Início Rápido (5 minutos)
- Guia de Contribuição para desenvolvedores
- Template CSV de exemplo para importação

#### Arquivos do Projeto
- Licença MIT com disclaimer legal
- .gitignore configurado para projetos Excel/VBA
- CONTRIBUTING.md com padrões de código
- CHANGELOG.md para rastreamento de versões

#### Funcionalidades Auxiliares
- Tratamento de erros em todas as funções principais
- Mensagens de feedback ao usuário
- Demonstração completa do sistema
- Exemplos de código para todas as operações

### Módulos VBA

1. **ModPrincipal.bas** (186 linhas)
   - Inicialização do sistema
   - Criação do dashboard
   - Exemplos de uso
   - Demonstração completa

2. **ModClientes.bas** (216 linhas)
   - Cadastro de clientes
   - Validação de CPF (algoritmo completo)
   - Busca e listagem
   - Inicialização da planilha

3. **ModImportacaoINSS.bas** (235 linhas)
   - Importação CSV
   - Cadastro manual de vínculos
   - Cálculo de tempos de contribuição
   - Inicialização de planilhas

4. **ModSimulacaoAposentadoria.bas** (298 linhas)
   - Simulação de 4 tipos de aposentadoria
   - Cálculos de elegibilidade
   - Verificação de requisitos
   - Geração de relatórios

### Recursos Técnicos

- **Total de linhas VBA:** 935
- **Documentação:** ~25.000 palavras
- **Exemplos:** 15+ snippets de código
- **Validações:** CPF, datas, duplicidade
- **Suporte:** CSV, Excel 2013+

## [Unreleased]

### Planejado para v1.1
- [ ] Conversão de tempo especial em comum (fator multiplicador)
- [ ] Interface gráfica com UserForms
- [ ] Exportação de relatórios em PDF
- [ ] Gráficos de evolução de contribuições
- [ ] Validação cruzada com CNIS

### Planejado para v1.2
- [ ] Projeções futuras (quando poderá se aposentar)
- [ ] Cálculo de benefício estimado
- [ ] Suporte para múltiplos vínculos simultâneos
- [ ] Comparação entre modalidades
- [ ] Timeline visual de contribuições

### Planejado para v2.0
- [ ] Integração com APIs do INSS
- [ ] Sistema multi-usuário
- [ ] Banco de dados externo
- [ ] Versão web

## Tipos de Mudanças

- `Adicionado` para novas funcionalidades
- `Modificado` para mudanças em funcionalidades existentes
- `Depreciado` para funcionalidades que serão removidas
- `Removido` para funcionalidades removidas
- `Corrigido` para correções de bugs
- `Segurança` para vulnerabilidades corrigidas

---

[1.0.0]: https://github.com/EduFBasso/Sist_Prev/releases/tag/v1.0.0
