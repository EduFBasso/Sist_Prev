# 📧 INSTRUÇÕES DE INSTALAÇÃO - ERP_PREV

## Para: Escritório de Advocacia

## De: Eduardo Figueiredo Basso

## Assunto: Sistema ERP_PREV - Versão de Testes

---

## 🎯 O QUE É ESTE SISTEMA?

Sistema para **cálculo automático de simulações previdenciárias** com:

- ✅ Importação automática de dados do CNIS (extrato INSS em PDF)
- ✅ Atualização automática de índices INPC (correção monetária)
- ✅ Cálculo de cenários de aposentadoria (Rápido, Equilibrado, Máximo)
- ✅ Interface completa em Excel com macros VBA

---

## 📦 CONTEÚDO DO E-MAIL

Você recebeu um arquivo compactado (.zip) contendo:

```
ERP_Prev_vX.Y.Z.zip (tamanho: ~25-30 MB)
```

**Dentro do .zip**:

```
📁 ERP_Prev_vX.Y.Z/
  ├─ 📄 erp_prev.xlsm           (arquivo principal Excel)
  ├─ 📄 MANUAL_USUARIO.pdf      (em breve)
  ├─ 📁 bin/
  │   ├─ atualizar_inpc.exe     (8 MB)
  │   ├─ converter_extrato_inss.exe (15 MB)
  │   └─ VERSAO.txt
  └─ 📁 saida/
      └─ (pasta vazia - será preenchida pelo sistema)
```

---

## 🚀 PASSO A PASSO - INSTALAÇÃO

### 1️⃣ DESCOMPACTAR ARQUIVO

**Opção A - Windows Explorer**:

1. Clicar com botão direito no arquivo `ERP_Prev_v1.0.zip`
2. Selecionar "Extrair tudo..."
3. Escolher destino: `C:\ERP_Prev\` (recomendado)
4. Clicar "Extrair"

**Opção B - 7-Zip / WinRAR**:

1. Abrir arquivo com 7-Zip ou WinRAR
2. Extrair para: `C:\ERP_Prev\`

**⚠️ IMPORTANTE**:

- Não deixar na pasta "Downloads" ou "Desktop"
- Caminho recomendado: `C:\ERP_Prev\` ou `C:\Projetos\ERP_Prev\`
- Evitar caminhos com acentos ou caracteres especiais
- Se receber/usar por pasta compartilhada (rede), prefira copiar para `C:\ERP_Prev\` antes de abrir

---

### 2️⃣ VERIFICAR CONTEÚDO

Abra a pasta `C:\ERP_Prev\ERP_Prev_v1.0\` e verifique:

✅ Arquivo principal: `erp_prev.xlsm` (alguns KB)
✅ Pasta `bin/` com 2 arquivos .exe (~23 MB total)
✅ Pasta `saida/` (vazia inicialmente)

Se algo estiver faltando, **não prossiga**. Entre em contato.

---

### 3️⃣ CONFIGURAR SEGURANÇA

#### A) Antivírus pode alertar (FALSO POSITIVO)

Os arquivos `.exe` são **seguros** mas podem acionar alerta de antivírus.

**Se o Windows Defender alertar**:

1. Abrir "Segurança do Windows"
2. Ir em "Proteção contra vírus e ameaças"
3. Clicar "Gerenciar configurações"
4. Em "Exclusões", clicar "Adicionar ou remover exclusões"
5. Adicionar pasta: `C:\ERP_Prev\`

**Outros antivírus** (Avast, AVG, Kaspersky):

- Adicionar pasta `C:\ERP_Prev\` nas exclusões
- Consultar suporte se necessário

#### B) Desbloquear arquivo Excel

1. Clicar com botão direito em `erp_prev.xlsm`
2. Selecionar "Propriedades"
3. Na aba "Geral", marcar ✅ "Desbloquear"
4. Clicar "Aplicar" → "OK"

#### C) (Recomendado) Adicionar Local Confiável no Excel

Se o arquivo estiver em pasta de rede, o Excel pode bloquear macros mesmo com "Habilitar Conteúdo".

1. Abrir Excel (em branco)
2. Arquivo → Opções → Central de Confiabilidade
3. Configurações da Central de Confiabilidade
4. Locais Confiáveis
5. Adicionar novo local: `C:\ERP_Prev\` (ou a pasta onde o sistema foi extraído)
6. Marcar “Subpastas deste local também são confiáveis”

---

### 4️⃣ ABRIR SISTEMA

1. **Duplo clique** em `erp_prev.xlsm`

2. Excel abrirá com **aviso de segurança** (amarelo no topo):

   ```
   AVISO DE SEGURANÇA: Macros foram desabilitadas
   ```

3. Clicar no botão **"Habilitar Conteúdo"**

4. Sistema carregará automaticamente

---

### 5️⃣ PRIMEIRO ACESSO

O sistema abrirá diretamente no **Formulário Principal**.

**Não precisa fazer nada agora!** Apenas confirmar que:

- ✅ Formulário abriu corretamente
- ✅ Não apareceu mensagem de erro
- ✅ Interface está visível e legível

Se tudo OK, **fechar o sistema** (botão X ou Sair).

---

## ✅ TESTE DE FUNCIONAMENTO

### Teste 1: Atualizar Índices INPC

1. Abrir `erp_prev.xlsm`
2. No formulário principal, clicar em **"Simulações"**
3. Clicar no botão **"Atualizar Índices"** (INPC)
4. Janela de terminal abrirá brevemente (5-10 segundos)
5. Mensagem de sucesso deve aparecer:
   ```
   Índices INPC atualizados com sucesso!
   Data: XX/XX/XXXX
   ```

Observação: no formulário principal existe a opção de **índice de correção** (INPC/SELIC).

- INPC: padrão
- SELIC: pode estar em validação e/ou não alterar valores dependendo da versão

**Se funcionar**: ✅ Sistema OK para uso

**Se der erro**: ❌ Anotar mensagem e enviar print

### Teste 2: Importar Dados (em breve)

(Instruções serão fornecidas após validação do Teste 1)

---

## 🆘 PROBLEMAS COMUNS

### ❌ "Arquivo não pode ser aberto"

**Causa**: Windows bloqueou arquivo baixado da internet

**Solução**: Seguir passo 3B (Desbloquear arquivo)

---

### ❌ "Macros foram desabilitadas permanentemente"

**Causa**: Configuração de segurança do Excel

**Solução**:

1. Abrir Excel (em branco)
2. Arquivo → Opções → Central de Confiabilidade
3. Configurações da Central de Confiabilidade
4. Configurações de Macro
5. Selecionar: "Habilitar todas as macros" (temporariamente)
6. Fechar Excel e abrir `erp_prev.xlsm` novamente

---

### ❌ "Arquivo inpc_fatores.csv não foi gerado"

**Causa**: Executável bloqueado ou problema de conectividade

**Soluções**:

1. Verificar se Windows Defender bloqueou (ver seção 3A)
2. Testar conexão com internet (sistema acessa BCB)
3. Executar manualmente:
   - Abrir pasta `C:\ERP_Prev\ERP_Prev_v1.0\bin\`
   - Duplo clique em `atualizar_inpc.exe`
   - Observar mensagens na janela preta

---

### ❌ Erro desconhecido

**Enviar para Eduardo**:

1. Print da tela do erro
2. Descrever o que estava fazendo
3. Versão do Windows (Win 10 / Win 11)
4. Versão do Excel (2016 / 2019 / 365)

---

## 📊 ESTRUTURA APÓS USO

Após usar o sistema, a estrutura ficará assim:

```
📁 ERP_Prev_v1.0/
  ├─ erp_prev.xlsm
  ├─ 📁 bin/
  │   └─ (executáveis)
  └─ 📁 saida/
      ├─ inpc_fatores.csv         (gerado na atualização INPC)
      └─ 📁 [NOME_CLIENTE]/       (gerado na importação)
          ├─ extrato_XXXXXX_dados_cliente.csv
          ├─ extrato_XXXXXX_remuneracoes.csv
          └─ extrato_XXXXXX_vinculos.csv
```

**Não deletar nada** desta estrutura!

---

## 🔄 ATUALIZAÇÕES FUTURAS

Este é um **sistema em testes**. Novas versões serão enviadas periodicamente.

**Ao receber nova versão**:

**Opção A - Sobrescrever (RECOMENDADO para testes)**:

1. Fechar Excel
2. Deletar pasta antiga: `C:\ERP_Prev\ERP_Prev_v1.0\`
3. Descompactar nova versão no mesmo lugar

**Opção B - Manter versões paralelas**:

1. Renomear pasta antiga: `ERP_Prev_v1.0_OLD`
2. Descompactar nova versão: `ERP_Prev_v1.1`
3. Comparar diferenças

**⚠️ BACKUP**: Se houver dados importantes, copiar pasta `saida/` antes de deletar.

---

## 📞 SUPORTE

**Desenvolvedor**: Eduardo Figueiredo Basso

**E-mail**: [edu_fabric@outlook.com]

**Disponibilidade**: Segunda a Sexta, 9h-18h

**Enviar**:

- Prints de tela dos erros
- Descrição detalhada do problema
- Arquivo `VERSAO.txt` da pasta `bin/`

---

## ✅ CHECKLIST DE INSTALAÇÃO

Marque conforme progride:

- [ ] Arquivo .zip descompactado em `C:\ERP_Prev\`
- [ ] Verificado conteúdo da pasta (erp_prev.xlsm + bin/ + saida/)
- [ ] Desbloquear arquivo Excel (Propriedades → Desbloquear)
- [ ] Adicionar exclusão no antivírus (se necessário)
- [ ] Abrir Excel e habilitar macros
- [ ] Formulário principal carregou corretamente
- [ ] Teste 1: Atualizar INPC funcionou ✅
- [ ] Arquivo `saida/inpc_fatores.csv` foi criado

---

**Versão deste documento**: 1.0  
**Data**: 13/01/2026  
**Sistema**: ERP_PREV v1.0.0
