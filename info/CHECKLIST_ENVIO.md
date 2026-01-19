# ✅ CHECKLIST DE VALIDAÇÃO PRÉ-ENVIO

## 📋 Sistema: ERP_PREV v1.0.0

## 👤 Responsável: Eduardo Figueiredo Basso

## 📅 Data: ******\_******

---

## 🎯 OBJETIVO

Garantir que o sistema está **100% funcional no Windows (Parallels)** antes de enviar ao escritório de advocacia para testes.

**IMPORTANTE**: Marcar **TODAS** as caixas antes do envio!

---

## 1️⃣ PREPARAÇÃO DO BUILD

### A) Geração dos Executáveis

- [ ] Python instalado no ambiente de desenvolvimento
- [ ] PyInstaller instalado: `pip install pyinstaller`
- [ ] Script `build_executaveis.py` executado com sucesso
- [ ] Pasta `bin/` contém 2 arquivos .exe:
  - [ ] `atualizar_inpc.exe` (~8 MB)
  - [ ] `converter_extrato_inss.exe` (~15 MB)
- [ ] Arquivo `bin/VERSAO.txt` criado com data/hora do build

**Comando para build**:

```bash
cd /Users/eduardofigueiredobasso/Documents/Sist_Prev
python build_executaveis.py
```

---

## 2️⃣ VALIDAÇÃO NO WINDOWS (PARALLELS)

### B) Ambiente Windows

- [ ] Windows iniciado (Parallels)
- [ ] Pasta compartilhada acessível do Windows
- [ ] Caminho visível: `C:\Users\...\Documents\Sist_Prev\`
- [ ] Excel instalado e funcional no Windows

### C) Teste dos Executáveis Standalone

**Teste Manual - atualizar_inpc.exe**:

- [ ] Abrir cmd no Windows
- [ ] Navegar: `cd C:\Users\...\Documents\Sist_Prev\bin`
- [ ] Executar: `atualizar_inpc.exe`
- [ ] Terminal exibe: "Conectando ao Banco Central..."
- [ ] Download completado: "~400 registros obtidos"
- [ ] Arquivo criado: `..\saida\inpc_fatores.csv`
- [ ] CSV contém competências e fatores (01/1979 até 12/2025)
- [ ] Fator base 12/2025 = 1.000000

**Teste Manual - converter_extrato_inss.exe** (se houver PDF):

- [ ] Copiar PDF de teste para pasta raiz
- [ ] Executar: `converter_extrato_inss.exe`
- [ ] Sistema solicita arquivo PDF
- [ ] Extração concluída sem erros
- [ ] 3 arquivos CSV criados na pasta `saida/[NOME_CLIENTE]/`

### D) Teste do Sistema Excel

**Abertura Inicial**:

- [ ] Abrir `erp_prev.xlsm` no Windows
- [ ] Windows mostra aviso de segurança (normal)
- [ ] Clicar "Habilitar Conteúdo"
- [ ] Formulário principal abre automaticamente
- [ ] Nenhuma mensagem de erro VBA

**Planilhas Base**:

- [ ] Planilha "Clientes" visível e acessível
- [ ] Planilha "Vinculos" visível e acessível
- [ ] Planilha "Remuneracoes" visível e acessível
- [ ] Planilha "Config_Regras" com 19+ parâmetros

**Cadastro de Cliente**:

- [ ] Abrir formulário "Cadastro"
- [ ] Todos os campos visíveis e editáveis
- [ ] Cadastrar cliente de teste: "Cliente Teste"
- [ ] Salvar com sucesso
- [ ] Cliente aparece na lista (busca)

**Simulações - Atualizar INPC**:

- [ ] Abrir formulário "Simulações"
- [ ] Selecionar cliente cadastrado na lista
- [ ] Clicar botão "Atualizar Índices INPC"
- [ ] Terminal abre e fecha automaticamente (5-10s)
- [ ] MsgBox de sucesso: "Índices INPC atualizados com sucesso!"
- [ ] Verificar: `Config_Regras` tem nova linha "Data_Atualizacao_INPC"
- [ ] Verificar: arquivo `saida/inpc_fatores.csv` existe

**Simulações - Calcular (com dados reais)**:

- [ ] Selecionar cliente com dados importados (ex: João Carlos)
- [ ] Clicar "Calcular Simulações"
- [ ] Tela exibe 3 cenários com valores
- [ ] Valores diferentes entre si (não todos iguais R$ 1.412)
- [ ] Valores realistas: R$ 2.000 - R$ 4.000 (com INPC aplicado)
- [ ] "Melhor regra" identificada corretamente
- [ ] Botão "Salvar" funciona (dados gravados em planilha)

**Importação de Dados** (se disponível):

- [ ] Formulário "Importar CNIS" acessível
- [ ] Testar importação com PDF de teste
- [ ] Dados cliente, vínculos e remunerações importados
- [ ] Valores nas planilhas correspondentes aos do PDF

---

## 3️⃣ VALIDAÇÃO DE DADOS

### E) Dados de Teste - João Carlos

- [ ] Cliente "João Carlos Eduardo Figueiredo Basso" existe
- [ ] Planilha "Vinculos": 13 vínculos cadastrados
- [ ] Planilha "Remuneracoes": 32 remunerações
- [ ] Competências no formato: MM/YYYY (ex: 04/1998)
- [ ] Valores originais: R$ 286 ~ R$ 1.616 (1995-2004)
- [ ] Após INPC: valores corrigidos para ~R$ 1.500 - R$ 5.000

### F) Cálculos Corretos

**SEM correção INPC** (deletar inpc_fatores.csv):

- [ ] Média calculada: ~R$ 1.000
- [ ] Todos cenários: R$ 1.412 (salário mínimo)

**COM correção INPC**:

- [ ] Média calculada: ~R$ 3.000 - R$ 3.500
- [ ] Cenário Rápido: R$ 2.100 - R$ 2.500
- [ ] Cenário Equilibrado: R$ 2.800 - R$ 3.200
- [ ] Cenário Máximo: R$ 3.500 - R$ 4.200

---

## 4️⃣ ESTRUTURA DE ARQUIVOS

### G) Estrutura para Envio

- [ ] Pasta raiz renomeada: `ERP_Prev_v1.0`
- [ ] Arquivo principal: `erp_prev.xlsm` presente
- [ ] Pasta `bin/` com 2 executáveis + VERSAO.txt
- [ ] Pasta `saida/` vazia (limpa para cliente)
- [ ] Remover arquivos temporários:
  - [ ] Deletar `__pycache__/`
  - [ ] Deletar `*.pyc`
  - [ ] Deletar `dist/` e `build/` (se existirem)
  - [ ] Deletar `*.spec` (se existirem)

### H) Arquivos de Documentação

- [ ] `README.md` atualizado com informações da v1.0
- [ ] `INSTALACAO_CLIENTE.md` incluído
- [ ] `MANUAL_USUARIO.pdf` (se disponível)
- [ ] `bin/VERSAO.txt` com build atual

---

## 5️⃣ SEGURANÇA E COMPATIBILIDADE

### I) Testes de Segurança

- [ ] Windows Defender: testado se bloqueia executáveis
- [ ] Se bloqueado: documentar solução (exclusão)
- [ ] Verificar propriedades do Excel: "Desbloquear" funciona
- [ ] Macros funcionam após desbloquear e habilitar

### J) Performance

- [ ] Abertura do Excel: < 5 segundos
- [ ] Atualizar INPC: < 15 segundos (com conexão internet)
- [ ] Calcular simulação: < 3 segundos
- [ ] Sistema não trava durante operações

---

## 6️⃣ COMPACTAÇÃO E ENVIO

### K) Preparar Pacote

- [ ] Fechar todos os arquivos Excel
- [ ] Fechar VBA Editor se aberto
- [ ] Limpar pasta `saida/` (deixar vazia)
- [ ] Compactar pasta: `ERP_Prev_v1.0.zip`
- [ ] Tamanho do .zip: ~25-30 MB
- [ ] Testar .zip: extrair em outra pasta e validar conteúdo

### L) E-mail de Envio

- [ ] Assunto: "ERP_PREV v1.0 - Sistema de Testes"
- [ ] Corpo: incluir instruções resumidas
- [ ] Anexo 1: `ERP_Prev_v1.0.zip`
- [ ] Anexo 2: `INSTALACAO_CLIENTE.md` (ou copiar no corpo)
- [ ] Destacar: "VERSÃO DE TESTES - aguardar feedback"
- [ ] Incluir contato para suporte
- [ ] Solicitar confirmação de recebimento

---

## 7️⃣ COMUNICAÇÃO COM CLIENTE

### M) Orientações Enviadas

- [ ] Passo a passo de descompactação
- [ ] Localização recomendada: `C:\ERP_Prev\`
- [ ] Instruções de desbloqueio (segurança Windows)
- [ ] Como habilitar macros no Excel
- [ ] Teste básico: "Atualizar INPC"
- [ ] O que fazer em caso de erro
- [ ] Contato para suporte (e-mail/telefone)

### N) Expectativas Definidas

- [ ] Informado: "Versão de TESTES"
- [ ] Informado: "Novas versões serão enviadas"
- [ ] Solicitado: feedback sobre funcionamento
- [ ] Solicitado: reportar qualquer erro
- [ ] Prazo de resposta: **\_\_** dias

---

## ✅ APROVAÇÃO FINAL

### Testes Obrigatórios Concluídos

- [ ] ✅ Build gerado com sucesso
- [ ] ✅ Executáveis testados manualmente (cmd)
- [ ] ✅ Excel abre sem erros
- [ ] ✅ Atualizar INPC funciona
- [ ] ✅ Simulação com João Carlos mostra valores corretos
- [ ] ✅ Estrutura de arquivos organizada
- [ ] ✅ Documentação incluída
- [ ] ✅ .zip criado e validado

### Assinaturas

**Desenvolvedor**: **************\_\_**************  
**Data/Hora**: **\_** / **\_** / **\_\_** às **\_**:**\_**

**Observações**:

```
_________________________________________________________________
_________________________________________________________________
_________________________________________________________________
```

---

## 📊 REGISTRO DE ENVIOS

| Versão | Data Envio     | Destinatário   | Arquivos Incluídos          | Status  |
| ------ | -------------- | -------------- | --------------------------- | ------- |
| v1.0.0 | **/**/\_\_\_\_ | Escritório XYZ | erp_prev.xlsm + bin/ + docs | Enviado |
| v1.0.1 | **/**/\_\_\_\_ | Escritório XYZ | Correções...                | -       |

---

## 🔄 VERSIONAMENTO

**Sistema de Versões**: X.Y.Z

- **X** (Major): Mudanças estruturais incompatíveis
- **Y** (Minor): Novas funcionalidades compatíveis
- **Z** (Patch): Correções de bugs

**Versão Atual**: 1.0.0  
**Próxima Versão Planejada**: 1.0.1 (correções após feedback)

---

**Documento criado em**: 13/01/2026  
**Última atualização**: 13/01/2026
