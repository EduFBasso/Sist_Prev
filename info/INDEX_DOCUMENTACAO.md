# 📁 ESTRUTURA DA DOCUMENTAÇÃO

**Última atualização:** 23/01/2026

---

## 📚 DOCUMENTAÇÃO ATIVA

### 🏗️ Arquitetura e Desenvolvimento
- **ARCHITECTURE.md** (raiz) - Arquitetura completa do conversor CNIS
- **PLANO_SIMPLIFICACAO.md** (raiz) - Roadmap de simplificação e modularização
- **README.md** - Visão geral do sistema

### 🔧 Extração CNIS
- **EXTRACAO_COMPLETA_ATUALIZADA.md** - Como funciona a extração (⚠️ atualizar com info de Facultativos)
- **EXTRACAO_REMUNERACOES.md** - Detalhes da extração de remunerações

### 🎨 Sistema VBA
- **MAPEAMENTO_CODIGO.md** - Estrutura do código VBA
- **MAPEAMENTO_FRMSIMULACOES.md** - Detalhes do formulário de simulações
- **MAPEAMENTO_SISTEMA_VBA.md** - Visão geral do sistema VBA
- **MAPEAMENTO_FLUXO_VINCULOS.md** - Fluxo de processamento de vínculos
- **MAPEAMENTO_PLANILHA_ADVOGADA.md** - Layout da planilha final
- **MOTOR_CALCULO_STATUS.md** - Status do motor de cálculo

### 📖 Guias e Regras
- **GUIA_RAPIDO.md** - Início rápido
- **GUIA_REGRAS_APOSENTADORIA.md** - Regras de aposentadoria EC 103/2019
- **PASSO_A_PASSO_SIMULACAO.md** - Como fazer uma simulação
- **MAPA_REGRAS_E_LAYOUT.md** - Mapeamento regras → layout

### ⚙️ Instalação e Configuração
- **INSTALACAO_CLIENTE.md** - Como instalar no cliente
- **INSTALACAO_PYTHON_WINDOWS.md** - Setup Python no Windows
- **INSTRUCOES_CONFIG_EMPRESA.md** - Configuração específica

### ✅ Validação e Testes
- **CHECKLIST_VALIDACAO.md** - Checklist de validação
- **EXEMPLOS_CALCULO_VALIDACAO.md** - Exemplos para validar cálculos
- **TESTE_INPC.md** - Testes de índice INPC
- **PROBLEMA_FORMATACAO_SELIC.md** - Problemas conhecidos SELIC

---

## 📦 DOCUMENTAÇÃO ARQUIVADA

Pasta: **info/_arquivados/**

Documentos que não refletem mais o estado atual do sistema:

- **VINCULOS_FACULTATIVOS_PENDENTE.md** - ✅ RESOLVIDO (23/01/2026)
  - Problema: Solicitava implementação de extração Facultativo
  - Status: Implementado no converter_extrato_inss.py
  - Ver: ARCHITECTURE.md para documentação atualizada

- **CORRECOES_IMPLEMENTADAS.md** - Histórico de correções antigas
- **STATUS_APOS_CORRECOES.md** - Status antigo do sistema
- **CHECKLIST_ENVIO.md** - Checklist antigo

---

## 🔄 DOCUMENTOS QUE PRECISAM ATUALIZAÇÃO

### Alta Prioridade
1. **EXTRACAO_COMPLETA_ATUALIZADA.md**
   - ⚠️ Não menciona implementação de Facultativos (Seq 11-13)
   - ⚠️ Não documenta dual detection (CLT vs Facultativo)
   - ✅ Sugestão: Mesclar com ARCHITECTURE.md ou adicionar referência

2. **EXTRACAO_REMUNERACOES.md**
   - ⚠️ Não menciona diferença entre "Remunerações" (CLT) e "Contribuições" (Facultativo)
   - ⚠️ Não documenta regex 5 campos

### Média Prioridade
3. **README.md**
   - ⚠️ Pode não refletir estrutura atual após limpeza
   - ✅ Adicionar referências a ARCHITECTURE.md e PLANO_SIMPLIFICACAO.md

---

## 📝 NOVAS ADIÇÕES (23/01/2026)

- ✅ **ARCHITECTURE.md** (raiz) - Documentação completa da arquitetura do conversor
- ✅ **PLANO_SIMPLIFICACAO.md** (raiz) - Roadmap de refatoração e modularização
- ✅ **info/INDEX_DOCUMENTACAO.md** (este arquivo) - Índice da documentação

---

## 🎯 RECOMENDAÇÕES

### Para Novos Desenvolvedores
1. Leia: **README.md** → **ARCHITECTURE.md** → **GUIA_RAPIDO.md**
2. Para extração CNIS: **ARCHITECTURE.md** (mais atualizado que EXTRACAO_*.md)
3. Para VBA: **MAPEAMENTO_SISTEMA_VBA.md** → **MAPEAMENTO_CODIGO.md**

### Para Manutenção
- Use **ARCHITECTURE.md** como fonte única de verdade para o conversor
- Documente mudanças no **PLANO_SIMPLIFICACAO.md**
- Arquive docs desatualizados em **_arquivados/**

### Para Usuários Finais
1. Instalação: **INSTALACAO_CLIENTE.md**
2. Uso: **GUIA_RAPIDO.md** → **PASSO_A_PASSO_SIMULACAO.md**
3. Dúvidas: **GUIA_REGRAS_APOSENTADORIA.md**

---

## 📊 ESTATÍSTICAS

- **Total documentos ativos:** 20 arquivos (.md)
- **Total arquivados:** 4 arquivos
- **Linhas de documentação:** ~4.500 linhas
- **Última grande atualização:** 23/01/2026 (Facultativos implementados)

---

## 🔗 LINKS RÁPIDOS

- Sistema Prev (raiz): `../`
- Backup desenvolvimento: `../_backup_desenvolvimento/`
- Scripts Python: `../` (5 scripts)
- Código VBA: `../Formularios/`, `../Modulos/`, `../Planilhas/`
