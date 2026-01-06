# Manual de Instalação - Sistema de Planejamento Previdenciário RGPS

## Visão Geral

Este sistema foi desenvolvido em Excel/VBA para auxiliar escritórios de advocacia previdenciária no planejamento de aposentadorias do RGPS (Regime Geral de Previdência Social).

## Requisitos

- Microsoft Excel 2013 ou superior
- Windows 7 ou superior / macOS com Excel
- Macros habilitadas

## Instalação

### Passo 1: Criar o Arquivo Excel

1. Abra o Microsoft Excel
2. Crie uma nova pasta de trabalho
3. Salve como "Sistema_Previdenciario.xlsm" (formato com macro habilitada)

### Passo 2: Habilitar a Guia Desenvolvedor

1. No Excel, clique em **Arquivo** → **Opções**
2. Selecione **Personalizar Faixa de Opções**
3. Marque a caixa **Desenvolvedor** no lado direito
4. Clique em **OK**

### Passo 3: Importar os Módulos VBA

1. Pressione `Alt + F11` para abrir o Editor VBA
2. No menu, clique em **Arquivo** → **Importar Arquivo...**
3. Navegue até a pasta `src/modules` deste repositório
4. Importe os seguintes arquivos na ordem:
   - `ModPrincipal.bas`
   - `ModClientes.bas`
   - `ModImportacaoINSS.bas`
   - `ModSimulacaoAposentadoria.bas`

### Passo 4: Configurar Segurança de Macros

1. No Excel, vá em **Arquivo** → **Opções** → **Central de Confiabilidade**
2. Clique em **Configurações da Central de Confiabilidade**
3. Selecione **Configurações de Macro**
4. Escolha **Habilitar todas as macros** (ou adicione o arquivo a um local confiável)
5. Marque **Confiar no acesso ao modelo de objeto do projeto VBA**
6. Clique em **OK**

### Passo 5: Inicializar o Sistema

1. Feche o Editor VBA (`Alt + Q`)
2. No Excel, pressione `Alt + F8` para abrir a lista de macros
3. Selecione `InicializarSistema`
4. Clique em **Executar**

O sistema criará automaticamente todas as planilhas necessárias:
- **Início**: Dashboard principal com instruções
- **Clientes**: Cadastro de clientes
- **Vínculos**: Registro de vínculos empregatícios
- **CalculoTempo**: Cálculos de tempo de contribuição
- **ResultadosSimulacao**: Resultados das simulações de aposentadoria

## Importação Alternativa via Copiar e Colar

Se preferir não importar arquivos, você pode copiar e colar o código manualmente:

1. Pressione `Alt + F11` para abrir o Editor VBA
2. No menu **Inserir**, selecione **Módulo**
3. Abra o arquivo `.bas` correspondente em um editor de texto
4. Copie todo o conteúdo (incluindo a primeira linha `Attribute VB_Name`)
5. Cole no módulo criado no Excel
6. Repita para cada módulo

## Verificação da Instalação

Após a instalação, você deve ver:

1. ✓ Cinco planilhas criadas
2. ✓ Quatro módulos VBA importados
3. ✓ Dashboard "Início" com instruções
4. ✓ Cabeçalhos formatados em todas as planilhas

## Solução de Problemas

### Erro: "Não é possível executar a macro"
- Verifique se as macros estão habilitadas
- Certifique-se de que salvou o arquivo como .xlsm

### Erro: "A planilha já existe"
- Exclua manualmente as planilhas existentes antes de executar `InicializarSistema`
- Ou feche e reabra o arquivo

### Erro ao importar módulos
- Certifique-se de que está importando arquivos .bas
- Verifique se a extensão do arquivo está visível no Windows

## Próximos Passos

Após a instalação, consulte:
- [Manual do Usuário](MANUAL_USUARIO.md) para instruções de uso
- [Guia de Importação](GUIA_IMPORTACAO.md) para importar dados do INSS

## Suporte

Para problemas ou dúvidas, consulte a documentação completa no repositório GitHub.
