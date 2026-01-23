# 🏗️ ARQUITETURA DO CONVERSOR CNIS

**Versão:** 2.0 (com suporte Facultativo)  
**Data:** 23/01/2026

---

## 📋 VISÃO GERAL

O `converter_extrato_inss.py` extrai dados do PDF do CNIS (Cadastro Nacional de Informações Sociais) e converte para CSVs estruturados.

**Entrada:** PDF do CNIS (8 páginas típico)  
**Saída:** 5 arquivos CSV

---

## 🎯 FORMATOS SUPORTADOS

### 1️⃣ CLT (Vínculos com Empregador)

**Exemplo no PDF:**
```
Matrícula do Tipo Filiado no
Seq. NIT          Código Emp.        Origem do Vínculo    Trabalhador...
2    125.37781.66-1  59.772.269/0001-39  DANTEK COM E IMP...  Empregado ou Agente

Remunerações
Competência  Remuneração  Indicadores
05/1996      286,25       13º SALÁRIO
06/1996      286,25
07/1996      286,25       MÚLTIPLOS VÍNCULOS
```

**Detecção:**
- ✅ Marcador: `"Matrícula do Tipo Filiado"`
- ✅ Possui campo: `"Código Emp."` (CNPJ)
- ✅ Seção dados: `"Remunerações"`

**Regex (3 campos):**
```python
r'(\d{2}/\d{4})\s+([\d.,]+)\s*([^\d/]*?)(?=\d{2}/\d{4}|$)'
#  ^^^^^^^^^^^^   ^^^^^^^^^^  ^^^^^^^^^^^
#  Competência    Remuneração Indicadores
```

**Características:**
- 3 campos por competência
- Até **3 competências por linha**
- CNPJ no campo `codigo_emp`

---

### 2️⃣ FACULTATIVO (Contribuinte Facultativo)

**Exemplo no PDF:**
```
Seq. NIT          Origem do Vínculo    Tipo Filiado    Vínculo      Data Início...
11   125.37781.66-1  RECOLHIMENTO      Facultativo     01/09/2019   30/10/2019

Contribuições
Competência  Data Pagto.  Contribuição  Salário Contrib.  Indicadores
09/2019      15/09/2019   200,00        1045,00           PREC-FACULTCONC
10/2019      15/10/2019   199,60        1045,00           PREC-FACULTCONC
```

**Detecção:**
- ✅ Marcador: `"Origem do Vínculo"` (sem "Matrícula do Tipo Filiado")
- ✅ Linha seguinte: padrão NIT (`\d{3}\.\d{5}\.\d{2}-\d`)
- ✅ Contém: `"RECOLHIMENTO"`
- ❌ NÃO possui: `"Código Emp."`
- ✅ Seção dados: `"Contribuições"`

**Regex (5 campos, captura 3):**
```python
r'(\d{2}/\d{4})\s+\d{2}/\d{2}/\d{4}\s+([\d.,]+)\s+([\d.,]+)\s*([^\d/]*?)(?=\d{2}/\d{4}|$)'
#  ^^^^^^^^^^^^   ^^^^^^^^^^^^^^^^   ^^^^^^^^^^  ^^^^^^^^^^  ^^^^^^^^^^^
#  Competência    Data Pagto (SKIP)  Contribuição Salário(SKIP) Indicadores
#  (captura)                         (captura)                  (captura)
```

**Características:**
- 5 campos totais, **captura 3** (ignora Data Pagto e Salário)
- Até **2 competências por linha** (menos denso)
- `codigo_emp = "FACULTATIVO"` (sem CNPJ)
- Campo `Contribuição` → salvo como `remuneracao` no CSV

---

## 🔄 FLUXO DE EXECUÇÃO

### 1. Inicialização
```python
pdf = pdfplumber.open(caminho_pdf)
seq_atual = None          # Mantém Seq entre páginas
codigo_emp_atual = None   # Mantém tipo (CNPJ ou "FACULTATIVO")
```

### 2. Loop por Página
Para cada página:

#### 2.1 Delimitar Zona Útil
```python
# Início: primeiro que encontrar
inicio_zona = texto.find("Relações Previdenciárias")
if inicio_zona == -1:
    inicio_zona = texto.find("Identificação do Filiado")

# Fim
fim_zona = texto.find("O INSS poderá rever")

zona_util = texto[inicio_zona:fim_zona]
```

**Por quê?** Elimina cabeçalhos, rodapés, QR codes → reduz falsos positivos

#### 2.2 Processar Continuação (se página > 1)
```python
if pagina > 1 and seq_atual and codigo_emp_atual:
    if codigo_emp_atual == "FACULTATIVO":
        # Procurar seção "Contribuições" na zona útil
        if "Contribuições" in zona_util:
            processar_contribuicoes_facultativo(...)
    else:  # CLT
        # Procurar valores soltos (mm/yyyy valor indicadores)
        for linha in zona_util:
            if "Matrícula" in linha: break
            regex_3_campos(linha) → adicionar em registros
```

**Por quê?** Vínculos longos continuam na próxima página. Estado `seq_atual` + `codigo_emp_atual` permite associar valores soltos ao vínculo correto.

#### 2.3 Buscar Próximo Vínculo (Dual Detection)

```python
pos = 0
while True:
    # Buscar CLT
    inicio_clt = zona_util.find("Matrícula do Tipo Filiado", pos)
    
    # Buscar Facultativo
    inicio_fac = -1
    idx_origem = zona_util.find("Origem do Vínculo", pos)
    if idx_origem != -1:
        proxima_linha = zona_util[idx_origem:idx_origem+200]
        if regex_NIT in proxima_linha and "RECOLHIMENTO" in proxima_linha:
            inicio_fac = idx_origem
    
    # Qual vem primeiro?
    if inicio_clt == -1 and inicio_fac == -1:
        break  # Fim da página
    
    if inicio_fac != -1 and (inicio_clt == -1 or inicio_fac < inicio_clt):
        processar_bloco_facultativo(...)
    else:
        processar_bloco_clt(...)
```

**Estratégia:** Busca ambos formatos e processa o que aparecer primeiro → suporta mistura CLT + Facultativo na mesma página.

---

### 3. Processamento por Tipo

#### 3.A. Bloco CLT

```python
# Extrair Seq e CNPJ
seq_match = re.search(r"Seq\.\s+.*?\n(\d+)", bloco)
cnpj_match = re.search(r"Código Emp\.\s+.*?\n\d+\s+[0-9.\-]+\s+([0-9./\-]+)", bloco)

seq_atual = seq_match.group(1)           # Ex: "2"
codigo_emp_atual = cnpj_match.group(1)   # Ex: "59.772.269/0001-39"

# Procurar seção "Remunerações"
inicio_remun = bloco.find("Remunerações")
secao_remun = bloco[inicio_remun:]

# Processar linhas
for linha in secao_remun.split('\n'):
    # Regex 3 campos: até 3 competências por linha
    padroes = re.findall(
        r'(\d{2}/\d{4})\s+([\d.,]+)\s*([^\d/]*?)(?=\d{2}/\d{4}|$)',
        linha
    )
    
    for competencia, remuneracao, indicadores in padroes:
        registros.append({
            "seq": seq_atual,
            "codigo_emp": codigo_emp_atual,
            "competencia": competencia,
            "remuneracao": remuneracao,
            "indicadores": indicadores,
            "pagina": pagina_idx
        })
```

#### 3.B. Bloco Facultativo

```python
# Extrair Seq da linha de dados (formato tabular)
linha_dados = zona_util[idx_origem:idx_origem+200].split('\n')[1]
seq_match = re.match(r'^\s*(\d+)\s+\d{3}\.\d{5}\.\d{2}-\d', linha_dados)

seq_atual = seq_match.group(1)     # Ex: "11"
codigo_emp_atual = "FACULTATIVO"   # SEM CNPJ

# Delimitar bloco (até próximo vínculo)
prox_clt = zona_util.find("Matrícula do Tipo Filiado", idx_origem + 50)
prox_fac = zona_util.find("Origem do Vínculo", idx_origem + 50)
fim_bloco = min([len(zona_util), prox_clt, prox_fac])

bloco_fac = zona_util[idx_origem:fim_bloco]

# Procurar seção "Contribuições"
if "Contribuições" in bloco_fac:
    secao_contrib = bloco_fac[bloco_fac.find("Contribuições"):]
    processar_contribuicoes_facultativo(secao_contrib, seq_atual, pagina_idx, registros)
```

**Função `processar_contribuicoes_facultativo`:**
```python
for linha in secao_contrib.split('\n'):
    # Regex 5 campos (captura 3): até 2 competências por linha
    padroes = re.findall(
        r'(\d{2}/\d{4})\s+\d{2}/\d{2}/\d{4}\s+([\d.,]+)\s+([\d.,]+)\s*([^\d/]*?)(?=\d{2}/\d{4}|$)',
        linha
    )
    
    for competencia, contribuicao, _salario_skip, indicadores in padroes:
        registros.append({
            "seq": seq_atual,
            "codigo_emp": "FACULTATIVO",         # Sem CNPJ
            "competencia": competencia,
            "remuneracao": contribuicao,         # Contribuição → campo remuneracao
            "indicadores": indicadores,
            "pagina": pagina_idx
        })
```

---

### 4. Avançar Posição

```python
# CLT
pos = fim_bloco  # Pula para depois do bloco processado

# Facultativo
pos = fim_bloco_fac  # Garante não re-encontrar o mesmo "Origem do Vínculo"
```

**Crítico:** Evita loop infinito ao garantir que `pos` sempre avança além do bloco processado.

---

## 📊 ESTRUTURA DE DADOS

### CSV Remunerações (`*_remuneracoes.csv`)

```csv
Pagina;Seq;CodigoEmp;Competencia;Remuneracao;Indicadores
1;1;56.528.946/0001-80;01/1995;286.25;
1;1;56.528.946/0001-80;02/1995;286.25;
6;10;03.782.845/0001-74;01/2005;1005.70;
7;11;FACULTATIVO;09/2019;200.00;PREC-FACULTCONC
7;11;FACULTATIVO;10/2019;199.60;PREC-FACULTCONC
7;12;FACULTATIVO;10/2024;282.40;PREC-FACULTCONC
```

**Campos:**
- `Pagina`: 1-indexed
- `Seq`: 1, 2, ..., 10 (CLT), 11, 12, 13 (Facultativo)
- `CodigoEmp`: CNPJ (CLT) ou "FACULTATIVO"
- `Competencia`: mm/yyyy
- `Remuneracao`: valor float (CLT: remuneração, Facultativo: contribuição)
- `Indicadores`: marcadores separados por espaço

---

## 🧩 COMPONENTES PRINCIPAIS

### `extrair_remuneracoes_texto()` 🔥 CORE
- **560 linhas**
- Loop principal de extração
- Gerencia estado seq_atual/codigo_emp_atual
- Dual detection CLT vs Facultativo
- Coordena processamento de continuações

### `processar_contribuicoes_facultativo()`
- **49 linhas**
- Processa seção "Contribuições"
- Regex 5 campos (captura 3)
- Até 2 competências por linha
- codigo_emp = "FACULTATIVO"

### `extrair_vinculos_texto()` + `parse_vinculo_texto()`
- **133 linhas**
- Extrai blocos "Matrícula do Tipo Filiado"
- Parse estruturado: Seq, NIT, CNPJ, Empresa, Datas

### `salvar_vinculos_estruturados()`
- **99 linhas**
- Consolida vínculos de tabelas + texto
- Deduplicação por (Seq, NIT, CNPJ)
- **Ordenação por Seq** (adicionado em 23/01/2026)
- Gera CSV estruturado

---

## 🎨 PADRÕES REGEX

### CLT - Remunerações (3 campos)
```regex
(\d{2}/\d{4})\s+([\d.,]+)\s*([^\d/]*?)(?=\d{2}/\d{4}|$)
```

**Exemplo linha:**
```
05/1996 286,25 13º SALÁRIO 06/1996 286,25 07/1996 286,25 MÚLTIPLOS VÍNCULOS
```

**Captura:**
1. `05/1996`, `286,25`, `13º SALÁRIO`
2. `06/1996`, `286,25`, ``
3. `07/1996`, `286,25`, `MÚLTIPLOS VÍNCULOS`

### Facultativo - Contribuições (5 campos → 3 capturados)
```regex
(\d{2}/\d{4})\s+\d{2}/\d{2}/\d{4}\s+([\d.,]+)\s+([\d.,]+)\s*([^\d/]*?)(?=\d{2}/\d{4}|$)
```

**Exemplo linha:**
```
09/2019 15/09/2019 200,00 1045,00 PREC-FACULTCONC 10/2019 15/10/2019 199,60 1045,00 PREC-FACULTCONC
```

**Captura:**
1. `09/2019`, `200,00`, `PREC-FACULTCONC` (ignora `15/09/2019` e `1045,00`)
2. `10/2019`, `199,60`, `PREC-FACULTCONC` (ignora `15/10/2019` e `1045,00`)

### NIT Pattern (detecção Facultativo)
```regex
\d{3}\.\d{5}\.\d{2}-\d
```

**Exemplo:** `125.37781.66-1`

---

## 🔍 CASOS ESPECIAIS

### 1. Continuação entre Páginas

**Problema:** Vínculo começa na página N, "Remunerações" continua na página N+1.

**Solução:**
```python
# Página 1: Detecta "Matrícula do Tipo Filiado" Seq 2
seq_atual = "2"
codigo_emp_atual = "59.772.269/0001-39"

# Página 2: Antes de buscar próximo vínculo
if pagina > 1 and seq_atual:
    # Processa valores soltos no topo
    for linha in zona_util:
        if "Matrícula" in linha: break  # Próximo vínculo
        regex_3_campos(linha) → associa a seq_atual
```

### 2. Múltiplas Competências por Linha

**CLT (até 3):**
```
01/2005 1005,70 02/2005 1005,70 03/2005 1005,70
```

**Facultativo (até 2):**
```
09/2019 15/09/2019 200,00 1045,00 PREC-FACULTCONC 10/2019 15/10/2019 199,60 1045,00
```

**Solução:** `re.findall()` captura todas ocorrências na linha.

### 3. Indicadores Multi-palavra

**Exemplo:** `PREC-FACULTCONC` ou `13º SALÁRIO` ou `MÚLTIPLOS VÍNCULOS`

**Regex:** `([^\d/]*?)` captura tudo até próxima data ou fim de linha.

### 4. Valores com Ponto e Vírgula

**Formato:** `1.005,70` (mil reais e setenta centavos)

**Limpeza:**
```python
remuneracao_limpa = remuneracao.replace('.', '').replace(',', '.')
# "1.005,70" → "1005.70" (float)
```

### 5. Loop Infinito (Facultativo)

**Problema:** `find("Origem do Vínculo")` pode re-encontrar o mesmo cabeçalho.

**Solução:**
```python
# Buscar próximo Facultativo a partir de DEPOIS da linha de dados
prox_fac = zona_util.find("Origem do Vínculo", linha_dados_fim + 50)
#                                                ^^^^^^^^^^^^^^^^^^^^
#                                                +50 chars para pular o cabeçalho atual
```

---

## 📈 PERFORMANCE

**João Carlos (8 páginas, 13 vínculos):**
- Tempo: ~2 segundos
- 178 remunerações extraídas
- 10 CLT (163 remunerações) + 3 Facultativos (15 contribuições)

**Complexidade:**
- `O(P × L)` onde P = páginas, L = linhas por página
- Regex `findall` é linear no tamanho da linha
- Dual detection adiciona overhead mínimo (~2 buscas por posição)

---

## 🚀 MELHORIAS FUTURAS

### Modularização (Sessão 3 do PLANO)
```
extrator/
├── extrator_clt.py          # processar_remuneracoes()
├── extrator_facultativo.py  # processar_contribuicoes_facultativo()
├── parser_vinculos.py       # extrair_vinculos_texto(), parse_vinculo_texto()
└── utils.py                 # limpar_remuneracao(), detectar_encoding()
```

### Novos Formatos
- Contribuinte Individual (CI)
- MEI (Microempreendedor Individual)
- Autônomo

### Validações
- Verificar NIT/CPF com dígito verificador
- Alertar sobre datas inválidas (ex: 32/13/2020)
- Detectar valores suspeitos (>R$ 1.000.000,00)

### Relatório de Extração
```
Extração Completa!
- 13 vínculos encontrados
  * 10 CLT (163 remunerações)
  * 3 Facultativos (15 contribuições)
- Páginas processadas: 8
- Alertas: 2 competências com concomitância
```

---

## 🧪 TESTES

### Caso de Teste: João Carlos

**Entrada:** `cnis/CNIS_JOAO_CARLOS.pdf`

**Esperado:**
```
Seq 1:   5 remunerações (01-05/1995)
Seq 2:  31 remunerações (05/1996-11/1998)
Seq 3:   2 remunerações
Seq 4:   6 remunerações
Seq 5:   3 remunerações
Seq 6:   4 remunerações
Seq 7:   7 remunerações
Seq 8:  38 remunerações
Seq 9:  21 remunerações
Seq 10: 46 remunerações (continuação página 6)
Seq 11:  2 contribuições (09-10/2019) ← Facultativo
Seq 12:  9 contribuições (10/2024-06/2025) ← Facultativo
Seq 13:  4 contribuições (08-11/2025) ← Facultativo

TOTAL: 178 remunerações
```

**Verificação:**
```powershell
$csv = Import-Csv teste_remuneracoes.csv -Delimiter ';'
$csv | Group-Object Seq | Select-Object Count,Name | Sort-Object {[int]$_.Name}
# Deve mostrar 13 grupos com counts acima
```

---

## 📚 REFERÊNCIAS

- **PDF João Carlos:** `cnis/CNIS_JOAO_CARLOS.pdf`
- **Plano de Simplificação:** `PLANO_SIMPLIFICACAO.md`
- **Info/Documentação:** `info/`
  - `MAPEAMENTO_CODIGO.md`
  - `MOTOR_CALCULO_STATUS.md`
  - `EXTRACAO_COMPLETA_ATUALIZADA.md`

---

**Última atualização:** 23/01/2026  
**Status:** ✅ Funcional (CLT + Facultativo)  
**Próxima fase:** Modularização (ver PLANO_SIMPLIFICACAO.md Sessão 3)
