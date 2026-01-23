# 📋 PLANO DE SIMPLIFICAÇÃO E REORGANIZAÇÃO

**Data:** 23/01/2026  
**Objetivo:** Simplificar código, remover arquivos de teste, reorganizar estrutura, preparar para modularização

---

## 🎯 ESTADO ATUAL

### ✅ Arquivos de Produção (MANTER)

#### Scripts Python Principais
1. **converter_extrato_inss.py** (714 linhas) ⭐ PRINCIPAL
   - Extração de CNIS PDF → CSV
   - Suporta CLT (Seq 1-10) e Facultativo (Seq 11-13)
   - Gera 5 arquivos: raw, vinculos_brutos, vinculos_estruturado, dados_cliente, remuneracoes

2. **atualizar_inpc.py** (234 linhas)
   - Atualiza índices INPC para cálculo de correção monetária

3. **atualizar_selic.py** (323 linhas) / **atualizar_selic_4390.py** (313 linhas)
   - Atualiza taxas SELIC
   - **DECISÃO PENDENTE:** Qual versão manter?

4. **sync_vba.py** (284 linhas)
   - Sincroniza código VBA com sistema

5. **build_executaveis.py** (225 linhas)
   - Gera executáveis standalone

#### Scripts de Utilidade (verificar necessidade)
- **calcular_tempo_joao.py** (184 linhas) - Script específico do caso João Carlos (deletar?)
- **verificar_4390.py** (24 linhas) - Verificação específica (deletar?)

#### Diretórios de Produção
- **cnis/** - PDFs e extrações (MANTER)
  - JOAO_CARLOS_EDUARDO_FIGUEIREDO_BASSO/ - Dados extraídos
  - CNIS_JOAO_CARLOS.pdf - PDF de teste
  
- **inss_simulacao/** - Scripts de geração de planilhas
  - criar_planilha_simples.py ⭐ Gera Excel com abas
  
- **Formularios/** - Código VBA (MANTER)
- **Modulos/** - Módulos VBA (MANTER)
- **Planilhas/** - Config VBA (MANTER)
- **info/** - Documentação (MANTER)
- **bin/** - Executáveis gerados
- **fat_inpc_selic/** - Dados de índices econômicos
- **entrada/** e **saida/** - Diretórios de processamento

---

## 🗑️ ARQUIVOS PARA BACKUP/REMOÇÃO

### Arquivos de Teste (33 arquivos, ~180KB)
**Mover para:** `_backup_desenvolvimento/testes/`

```
teste_output.csv + 4 variações (_dados_cliente, _remuneracoes, _vinculos_*)
teste2.csv + 4 variações
teste3.csv + 4 variações  ✅ 163 remunerações (CLT completo)
teste4.csv + 4 variações  ⚠️ 165 remunerações (tentativa Facultativo)
teste5.csv + 4 variações  ✅ 178 remunerações (Facultativo funcionando)
teste6.csv + 4 variações  ✅ 178 remunerações (ordem corrigida) → ÚLTIMA VERSÃO BOA
```

**AÇÃO:** Mover teste6_* para cnis/JOAO_CARLOS.../cnis_* (substituir versão antiga)

### Scripts de Debug (10 arquivos, ~50KB)
**Mover para:** `_backup_desenvolvimento/scripts_debug/`

```
teste_facultativo.py       - Debug formato Facultativo
teste_regex_contrib.py     - Test regex contribuições
teste_regex.py             - Test regex básico
buscar_1997_1998.py        - Debug datas específicas
debug_blocos_pag1.py       - Debug página 1
debug_pagina3.py           - Debug página 3
debug_seq11.py             - Debug Seq 11
testar_seq2.py             - Debug Seq 2
ver_pagina2.py             - Debug página 2
detector_formato_cnis.py   - Detecta formato (pode ser útil)
```

### Versões Antigas de Código
**Mover para:** `_backup_desenvolvimento/versoes_antigas/`

```
converter_extrato_inss_backup.py  (1009 linhas) - Backup manual
converter_extrato_inss_old.py     (1009 linhas) - Versão antiga git
```

---

## 📁 ESTRUTURA PROPOSTA

```
Sist_Prev/
├── 📄 converter_extrato_inss.py         ⭐ PRINCIPAL
├── 📄 atualizar_inpc.py
├── 📄 atualizar_selic.py                (decidir qual versão)
├── 📄 sync_vba.py
├── 📄 build_executaveis.py
├── 📄 requirements.txt
├── 📄 README.txt
├── 📄 CHANGELOG.txt
│
├── 📁 cnis/                             # PDFs e extrações
│   ├── JOAO_CARLOS_EDUARDO_FIGUEIREDO_BASSO/
│   │   ├── cnis_raw.csv
│   │   ├── cnis_dados_cliente.csv
│   │   ├── cnis_vinculos_estruturado.csv
│   │   ├── cnis_remuneracoes.csv        ✅ 178 remunerações
│   │   └── cnis_validacao.txt
│   └── CNIS_JOAO_CARLOS.pdf
│
├── 📁 inss_simulacao/                   # Geração de planilhas Excel
│   ├── criar_planilha_simples.py        ⭐ PRINCIPAL
│   └── planilha_simples_YYYYMMDD_HHMMSS.xlsx
│
├── 📁 Formularios/                      # VBA Forms
├── 📁 Modulos/                          # VBA Modules
├── 📁 Planilhas/                        # VBA Config
├── 📁 info/                             # Documentação
├── 📁 bin/                              # Executáveis
├── 📁 fat_inpc_selic/                   # Dados econômicos
├── 📁 entrada/                          # Input folder
├── 📁 saida/                            # Output folder
│
└── 📁 _backup_desenvolvimento/          🆕 NOVA PASTA
    ├── testes/                          # 33 arquivos teste*.csv
    ├── scripts_debug/                   # 10 scripts debug
    └── versoes_antigas/                 # converter_extrato_inss_*.py
```

---

## 🔧 PLANO DE AÇÃO - FASE 1: LIMPEZA

### Etapa 1.1: Criar estrutura de backup
```powershell
New-Item -ItemType Directory -Path "_backup_desenvolvimento/testes" -Force
New-Item -ItemType Directory -Path "_backup_desenvolvimento/scripts_debug" -Force
New-Item -ItemType Directory -Path "_backup_desenvolvimento/versoes_antigas" -Force
```

### Etapa 1.2: Mover arquivos de teste
```powershell
Move-Item teste*.csv "_backup_desenvolvimento/testes/"
Move-Item teste*.py "_backup_desenvolvimento/scripts_debug/"
Move-Item *debug*.py "_backup_desenvolvimento/scripts_debug/"
Move-Item buscar_*.py "_backup_desenvolvimento/scripts_debug/"
Move-Item ver_*.py "_backup_desenvolvimento/scripts_debug/"
Move-Item testar_*.py "_backup_desenvolvimento/scripts_debug/"
Move-Item detector_formato_cnis.py "_backup_desenvolvimento/scripts_debug/"
```

### Etapa 1.3: Mover versões antigas
```powershell
Move-Item converter_extrato_inss_backup.py "_backup_desenvolvimento/versoes_antigas/"
Move-Item converter_extrato_inss_old.py "_backup_desenvolvimento/versoes_antigas/"
```

### Etapa 1.4: Atualizar dados finais do João Carlos
```powershell
Copy-Item teste6_dados_cliente.csv cnis/JOAO_CARLOS_EDUARDO_FIGUEIREDO_BASSO/cnis_dados_cliente.csv -Force
Copy-Item teste6_vinculos_estruturado.csv cnis/JOAO_CARLOS_EDUARDO_FIGUEIREDO_BASSO/cnis_vinculos_estruturado.csv -Force
Copy-Item teste6_remuneracoes.csv cnis/JOAO_CARLOS_EDUARDO_FIGUEIREDO_BASSO/cnis_remuneracoes.csv -Force
```

### Etapa 1.5: Decidir sobre scripts específicos
- [ ] **calcular_tempo_joao.py** - Deletar ou mover para _backup?
- [ ] **verificar_4390.py** - Deletar ou mover para _backup?
- [ ] **atualizar_selic.py vs atualizar_selic_4390.py** - Qual manter?

---

## 🔧 PLANO DE AÇÃO - FASE 2: MODULARIZAÇÃO

### 📦 Estrutura Modular Proposta

```
Sist_Prev/
├── 📄 converter_extrato_inss.py         # CLI principal (simplificado)
│
└── 📁 extrator/                         🆕 NOVO MÓDULO
    ├── __init__.py
    ├── core.py                          # Funções principais
    ├── extrator_clt.py                  # Extração CLT (Remunerações)
    ├── extrator_facultativo.py          # Extração Facultativo (Contribuições)
    ├── parser_vinculos.py               # Parse de vínculos
    └── utils.py                         # Utilitários (encoding, etc)
```

### 2.1 Separação de Responsabilidades

**converter_extrato_inss.py** (SIMPLIFICADO - ~100 linhas)
```python
# Apenas CLI e orquestração
- parse_args()
- main()
- chamar funções do módulo extrator.*
```

**extrator/core.py**
```python
# Funções centrais
- extrair_dados_cliente()
- coordenar_extracao_paginas()
- salvar_csv_raw()
```

**extrator/extrator_clt.py**
```python
# Específico para CLT (3 campos)
- processar_remuneracoes()
- regex: mm/yyyy valor indicadores
- detectar: "Matrícula do Tipo Filiado" + "Código Emp." presente
```

**extrator/extrator_facultativo.py**
```python
# Específico para Facultativo (5 campos)
- processar_contribuicoes_facultativo()
- regex: mm/yyyy dd/mm/yyyy contribuição salário indicadores
- detectar: "Origem do Vínculo" + NIT pattern + "RECOLHIMENTO" + SEM "Código Emp."
```

**extrator/parser_vinculos.py**
```python
# Parse de vínculos (tabela + texto)
- extrair_vinculos_texto()
- parse_vinculo_texto()
- salvar_vinculos_estruturados()
```

**extrator/utils.py**
```python
# Funções auxiliares
- detectar_encoding()
- limpar_remuneracao()
- validar_valor()
```

---

## 📊 ANÁLISE DO CÓDIGO ATUAL

### converter_extrato_inss.py (714 linhas)

**Estrutura Atual:**
```
Linhas 1-80:    Imports, constantes, encoding
Linhas 83-150:  extrair_vinculos_texto(), parse_vinculo_texto()
Linhas 152-216: parse_vinculo_texto() continuação
Linhas 216-315: salvar_vinculos_estruturados()
Linhas 317-394: extrair_dados_cliente(), processar_remuneracoes()
Linhas 394-443: processar_contribuicoes_facultativo() 🆕 FACULTATIVO
Linhas 446-651: extrair_remuneracoes_texto() 🔥 FUNÇÃO PRINCIPAL
Linhas 653-714: main(), CLI, orquestração
```

**Blocos Identificados:**

1. **BLOCO CLT** (Remunerações - 3 campos)
   - Linhas 317-392: `processar_remuneracoes()`
   - Marcador: "Matrícula do Tipo Filiado" + "Código Emp."
   - Seção: "Remunerações"
   - Regex: `(\d{2}/\d{4})\s+([\d.,]+)\s*([^\d/]*?)(?=\d{2}/\d{4}|$)`
   - Formato: Competência | Remuneração | Indicadores
   - Até 3 competências por linha

2. **BLOCO FACULTATIVO** (Contribuições - 5 campos)
   - Linhas 394-443: `processar_contribuicoes_facultativo()`
   - Marcador: "Origem do Vínculo" + NIT + "RECOLHIMENTO"
   - Seção: "Contribuições"
   - Regex: `(\d{2}/\d{4})\s+\d{2}/\d{2}/\d{4}\s+([\d.,]+)\s+([\d.,]+)\s*([^\d/]*?)(?=\d{2}/\d{4}|$)`
   - Formato: Competência | Data Pagto | Contribuição | Salário | Indicadores
   - Ignora: Data Pagto, Salário
   - Captura: Competência, Contribuição (como remuneracao), Indicadores
   - Até 2 competências por linha

3. **BLOCO VÍNCULOS** (Parse estruturado)
   - Linhas 83-150: `extrair_vinculos_texto()`
   - Linhas 152-216: `parse_vinculo_texto()`
   - Linhas 216-315: `salvar_vinculos_estruturados()`
   - Extração: Tabelas + Texto direto
   - Deduplicação: (Seq, NIT, CNPJ)
   - Ordenação: Seq numérico 🆕 ADICIONADO

4. **BLOCO COORDENAÇÃO** (Loop principal)
   - Linhas 446-651: `extrair_remuneracoes_texto()`
   - Zona útil: Entre "Relações Previdenciárias" e "O INSS poderá rever"
   - Continuação: page > 1 + seq_atual + codigo_emp_atual
   - Dual detection: CLT OR Facultativo
   - Estado: seq_atual, codigo_emp_atual (mantido entre páginas)

---

## 🎯 BENEFÍCIOS DA MODULARIZAÇÃO

### ✅ Vantagens

1. **Separação de Conceitos**
   - CLT e Facultativo em arquivos separados
   - Mais fácil entender cada formato
   - Reduz acoplamento

2. **Testabilidade**
   - Testar CLT sem Facultativo
   - Testar Facultativo isoladamente
   - Unit tests por módulo

3. **Manutenção**
   - Bug em CLT? Editar apenas extrator_clt.py
   - Novo formato? Adicionar extrator_autonomo.py sem quebrar existentes
   - Código menor = menos bugs

4. **Reuso**
   - Funções utils compartilhadas
   - Parser de vínculos independente
   - Pode ser usado por outros scripts

5. **Documentação**
   - Cada módulo com docstring específico
   - README por responsabilidade
   - Exemplos isolados

### ⚠️ Cuidados

1. **Não quebrar interface CLI**
   - Manter `converter_extrato_inss.py [PDF] [OUTPUT.csv]`
   - Imports internos transparentes

2. **Performance**
   - Não duplicar leitura de PDF
   - Passar objetos pdfplumber entre funções
   - Cache de zona útil se necessário

3. **Compatibilidade**
   - Formato CSV igual
   - Ordem de colunas preservada
   - Scripts downstream (criar_planilha_simples.py) não quebram

---

## 🚀 CRONOGRAMA SUGERIDO

### Sessão 1: Limpeza (30 min)
- [ ] Criar pastas _backup_desenvolvimento/
- [ ] Mover arquivos teste*
- [ ] Mover scripts debug
- [ ] Mover versões antigas
- [ ] Atualizar dados finais João Carlos
- [ ] Commit: "chore: Limpar arquivos de desenvolvimento"

### Sessão 2: Documentação do Código Atual (1h)
- [ ] Adicionar docstrings detalhados
- [ ] Comentar blocos CLT vs Facultativo
- [ ] Documentar regex patterns
- [ ] Criar ARCHITECTURE.md explicando fluxo
- [ ] Commit: "docs: Documentar arquitetura do extrator"

### Sessão 3: Extração de Funções (1h30)
- [ ] Criar módulo extrator/__init__.py
- [ ] Mover processar_remuneracoes() → extrator_clt.py
- [ ] Mover processar_contribuicoes_facultativo() → extrator_facultativo.py
- [ ] Mover funções de vínculo → parser_vinculos.py
- [ ] Testar: python converter_extrato_inss.py cnis/CNIS_JOAO_CARLOS.pdf teste_refactor.csv
- [ ] Verificar: 178 remunerações preservadas
- [ ] Commit: "refactor: Modularizar extratores CLT e Facultativo"

### Sessão 4: Simplificação CLI (30 min)
- [ ] Reduzir converter_extrato_inss.py para ~150 linhas
- [ ] Manter apenas orquestração e CLI
- [ ] Imports do módulo extrator.*
- [ ] Testar novamente
- [ ] Commit: "refactor: Simplificar CLI principal"

### Sessão 5: Testes e Validação (1h)
- [ ] Testar com PDF João Carlos (178 remunerações)
- [ ] Verificar ordem Seq 1-10
- [ ] Verificar Seq 11-13 Facultativo
- [ ] Gerar planilha Excel
- [ ] Validar todas as abas
- [ ] Commit: "test: Validar extração completa pós-refatoração"

### Sessão 6: Adicionar Facultativos na Aba Vínculos (1h)
- [ ] Analisar estrutura aba Vínculos
- [ ] Definir formato para Facultativo (sem CNPJ, tipo "Facultativo")
- [ ] Criar função para extrair vínculos Facultativo
- [ ] Integrar no salvar_vinculos_estruturados()
- [ ] Testar: Seq 11-13 devem aparecer na aba
- [ ] Commit: "feat: Adicionar vínculos Facultativos na planilha"

---

## 📝 PRÓXIMOS PASSOS (PÓS-SIMPLIFICAÇÃO)

### Melhorias Futuras

1. **Validação de Dados**
   - Verificar datas válidas
   - Alertar sobre valores suspeitos
   - Validar NIT/CPF com dígito verificador

2. **Relatório de Extração**
   - Gerar log detalhado
   - Estatísticas: N vínculos, N remunerações por Seq
   - Alertas de qualidade

3. **Suporte a Mais Formatos**
   - Contribuinte Individual (CI)
   - MEI (Microempreendedor Individual)
   - Autônomo
   - Cada um com extrator específico

4. **Interface Gráfica**
   - GUI simples para arrastar PDF
   - Visualizar prévia dos dados
   - Gerar Excel diretamente

5. **Testes Automatizados**
   - pytest para cada módulo
   - Fixtures com PDFs sintéticos
   - CI/CD com GitHub Actions

---

## 📋 CHECKLIST FINAL

### Antes de Comitar
- [ ] Todos os testes passam (178 remunerações)
- [ ] Planilha Excel gerada corretamente
- [ ] Sem arquivos teste* no diretório raiz
- [ ] README.txt atualizado
- [ ] CHANGELOG.txt atualizado com versão
- [ ] Git status limpo (exceto _backup_desenvolvimento/)

### Git Commits Sugeridos
```bash
git add _backup_desenvolvimento/
git commit -m "chore: Criar estrutura de backup para arquivos de desenvolvimento"

git rm teste*.csv teste*.py *debug*.py
git commit -m "chore: Remover arquivos de teste e debug"

git add converter_extrato_inss.py
git commit -m "fix: Ordenar vínculos por Seq + adicionar extração Facultativo"

git add inss_simulacao/criar_planilha_simples.py
git commit -m "fix: Ordenar vínculos na planilha Excel"

git add extrator/
git commit -m "refactor: Modularizar extratores CLT e Facultativo"

git push origin main
```

---

## 🎓 LIÇÕES APRENDIDAS

### Durante o Desenvolvimento

1. **Git Restore Salvou o Dia**
   - Sempre manter commits funcionais
   - Testar antes de comitar grandes mudanças

2. **Abordagem Incremental**
   - Pequenas mudanças + teste = sucesso
   - Big bang refactor = desastre

3. **Testar com Dados Reais**
   - PDF João Carlos foi essencial
   - 178 remunerações = métrica de sucesso

4. **Documentação Durante Desenvolvimento**
   - Arquivos info/*.md ajudaram muito
   - Mapear código enquanto desenvolve

### Arquitetura

1. **Detecção por Estrutura é Robusta**
   - "Código Emp." presente = CLT
   - "RECOLHIMENTO" + sem "Código Emp." = Facultativo
   - Melhor que hardcoded Seq numbers

2. **Estado Entre Páginas é Crucial**
   - seq_atual + codigo_emp_atual
   - Permite continuação em páginas seguintes

3. **Zona Útil Elimina Ruído**
   - Focar entre marcadores
   - Ignora header/footer
   - Reduz falsos positivos

---

**FIM DO PLANO**

*Bom almoço! 🍽️ Quando voltar, podemos executar qualquer fase deste plano.*
