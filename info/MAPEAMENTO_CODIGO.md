# 🗺️ MAPEAMENTO COMPLETO DO CÓDIGO - ERP_PREV v1.0

## 📋 Objetivo
Documentar TODOS os arquivos, funções e dependências do sistema para garantir funcionamento correto antes do envio ao cliente.

---

## 📂 ESTRUTURA DE ARQUIVOS

### Arquivos Python (.py)

| Arquivo | Localização | Função | Status | Dependências |
|---------|-------------|--------|--------|--------------|
| `atualizar_inpc.py` | Raiz | Baixa série 188 INPC do BCB, calcula fatores de correção | ✅ Pronto | urllib, json, os, datetime |
| `converter_extrato_inss.py` | Raiz | Converte PDF CNIS para CSV | ✅ Pronto | PyMuPDF, re, csv |
| `build_executaveis.py` | Raiz | Gera executáveis .exe com PyInstaller | ✅ Pronto | pyinstaller, shutil, subprocess |

### Arquivos VBA (.bas)

| Módulo | Arquivo | Funções Principais | Linhas | Status |
|--------|---------|-------------------|--------|--------|
| modSimulacoes | Modulos/modSimulacoes.bas | AtualizarIndicesINPC, AplicarCorrecaoINPC, CalcularMediaSalarios | ~1278 | ✅ Ajustado |
| modImportacao | Modulos/modImportacao.bas | ImportarDadosCNIS, ProcessarVinculos | - | ⚠️ Verificar |
| modDB | Modulos/modDB.bas | GetParametro, SetParametro | - | ✅ Pronto |
| modCadastro | Modulos/modCadastro.bas | SalvarCliente, BuscarCliente | - | ✅ Pronto |
| modVinculos | Modulos/modVinculos.bas | GerenciarVinculos | - | ✅ Pronto |
| modUtil | Modulos/modUtil.bas | Funções auxiliares | - | ✅ Pronto |

### Formulários VBA (.bas)

| Formulário | Arquivo | Função | Status |
|------------|---------|--------|--------|
| frmPrincipal | Formularios/frmPrincipal.bas | Menu principal do sistema | ✅ Pronto |
| frmCadastro | Formularios/frmCadastro.bas | Cadastro de clientes | ✅ Pronto |
| frmSimulacoes | Formularios/frmSimulacoes.bas | Cálculo de simulações | ✅ Pronto |
| frmVinculos | Formularios/frmVinculos.bas | Gestão de vínculos | ✅ Pronto |
| frmBusca | Formularios/frmBusca.bas | Busca de clientes | ✅ Pronto |

### Planilhas Excel

| Planilha | Colunas | Função | Status |
|----------|---------|--------|--------|
| Clientes | ID, Nome, CPF, DN, Sexo, etc. | Armazena cadastro de clientes | ✅ OK |
| Vinculos | ID, ID_Cliente, Empresa, Tipo, Início, Fim | Armazena vínculos empregatícios | ✅ OK |
| Remuneracoes | ID, ID_Cliente, Competência, Valor | Armazena salários históricos | ✅ OK |
| Config_Regras | Parametro, Valor, Descricao, Data | Parâmetros do sistema | ✅ OK |

---

## 🔗 FLUXO DE INTEGRAÇÃO PYTHON-VBA

### 1️⃣ Atualização INPC (BCB → CSV → VBA)

```
📊 FLUXO COMPLETO:

[Usuário] → Clica "Atualizar Índices INPC" em frmSimulacoes
   ↓
[VBA] modSimulacoes.AtualizarIndicesINPC()
   ↓
   Detecta: bin/atualizar_inpc.exe OU atualizar_inpc.py
   ↓
   Shell(comando) → Executa processo externo
   ↓
[Python] atualizar_inpc.py inicia
   ↓
   urlopen("https://api.bcb.gov.br/dados/serie/bcdata.sgs.188/dados?formato=json")
   ↓
   Recebe ~400 registros (01/1979 até 12/2025)
   ↓
   Calcula fatores: Fator = INPC_base(12/2025) / INPC_mes
   ↓
   Salva: saida/inpc_fatores.csv (Competencia;Fator)
   ↓
[VBA] Verifica criação do arquivo CSV
   ↓
   SetParametro("Data_Atualizacao_INPC", Date)
   ↓
   MsgBox("Sucesso!")
   ↓
[Usuário] → Vê confirmação
```

### 2️⃣ Cálculo de Simulação (VBA → CSV → Cálculo)

```
📊 FLUXO COMPLETO:

[Usuário] → Seleciona cliente + Clica "Calcular Simulações"
   ↓
[VBA] frmSimulacoes.cmdCalcular_Click()
   ↓
   modSimulacoes.CalcularCenarioRapido(ID_Cliente)
   ↓
   CalcularValorBeneficio(anos, coeficiente, ID_Cliente)
   ↓
   CalcularMediaSalarios(ID_Cliente)
      ↓
      Lê planilha "Remuneracoes": 32 salários
      ↓
      Para cada salário:
         valor_original = Cells(i, 5)
         competencia = Cells(i, 4)  "04/1998"
         ↓
         AplicarCorrecaoINPC(valor_original, competencia)
            ↓
            [CACHE] Lê saida/inpc_fatores.csv (1x por dia)
            ↓
            Dictionary: {"04/1998": 5.234567}
            ↓
            valor_corrigido = 286 * 5.234567 = ~1.497
            ↓
            Retorna valor_corrigido
      ↓
      Ordena 32 valores decrescente
      ↓
      Calcula média dos 80% maiores (26 salários)
      ↓
      Retorna: ~R$ 3.500
   ↓
   Aplica coeficiente: 60% = R$ 2.100
   ↓
   Aplica limites (mínimo/teto)
   ↓
   Retorna valor final ao formulário
   ↓
[Usuário] → Vê valor na tela
```

---

## 🔍 FUNÇÕES CRÍTICAS DETALHADAS

### modSimulacoes.AtualizarIndicesINPC()

**Localização**: Modulos/modSimulacoes.bas, linha ~1139

**Objetivo**: Executar atualização INPC via Python/exe

**Lógica**:
```vb
1. caminhoExe = ThisWorkbook.Path & "\bin\atualizar_inpc.exe"
2. caminhoScript = ThisWorkbook.Path & "\atualizar_inpc.py"

3. IF Dir(caminhoExe) <> "" THEN
     comando = """" & caminhoExe & """"  ' Produção
   ELSEIF Dir(caminhoScript) <> "" THEN
     comando = "python """ & caminhoScript & """"  ' Desenvolvimento
   ELSE
     ERRO: Nenhum executável encontrado
   END IF

4. Shell(comando, vbNormalFocus)
5. Application.Wait (5 segundos)
6. Verificar: saida/inpc_fatores.csv existe
7. SetParametro("Data_Atualizacao_INPC", Date)
8. MsgBox sucesso
```

**Dependências**:
- ✅ modDB.SetParametro()
- ✅ bin/atualizar_inpc.exe (produção)
- ⚠️ atualizar_inpc.py + Python (desenvolvimento)

**Testado**: ⏳ Pendente teste no Windows

---

### modSimulacoes.AplicarCorrecaoINPC()

**Localização**: Modulos/modSimulacoes.bas, linha ~1220

**Objetivo**: Multiplicar valor por fator INPC da competência

**Lógica**:
```vb
1. [STATIC] Dictionary fatoresINPC (cache)
2. [STATIC] Date ultimaLeitura

3. IF fatoresINPC Is Nothing OR DateDiff > 0 THEN
     ' Carregar CSV
     Open "saida/inpc_fatores.csv"
     FOR EACH linha:
       partes = Split(linha, ";")
       competencia = partes(0)  ' "04/1998"
       fator = CDbl(partes(1))   ' 5.234567
       fatoresINPC(competencia) = fator
     NEXT
     ultimaLeitura = Date
   END IF

4. IF fatoresINPC.Exists(competencia) THEN
     RETURN valor * fatoresINPC(competencia)
   ELSE
     RETURN valor  ' Sem correção se não encontrado
   END IF
```

**Performance**: 
- ✅ CSV lido 1x por dia (cache static)
- ✅ ~400 fatores em memória (~10 KB)
- ✅ Lookup O(1) via Dictionary

**Testado**: ⏳ Pendente teste com dados reais

---

### modSimulacoes.CalcularMediaSalarios()

**Localização**: Modulos/modSimulacoes.bas, linha ~1056

**Objetivo**: Calcular média dos 80% maiores salários CORRIGIDOS

**MODIFICAÇÃO CRÍTICA** (linha ~1084):
```vb
' ANTES (sem INPC):
salarios(n) = ws.Cells(i, 5).Value

' DEPOIS (com INPC):
valor = ws.Cells(i, 5).Value          ' R$ 286
competencia = ws.Cells(i, 4).Value    ' "04/1998"
valorCorrigido = AplicarCorrecaoINPC(valor, competencia)  ' R$ 1.497
salarios(n) = valorCorrigido
```

**Fluxo**:
1. Ler planilha "Remuneracoes", filtrar ID_Cliente
2. Para cada remuneração: aplicar correção INPC
3. Coletar array de salários corrigidos
4. Ordenar decrescente (bubble sort)
5. Calcular média dos 80% maiores
6. Retornar valor

**Testado**: ⏳ João Carlos deve retornar ~R$ 3.500 (em vez de R$ 1.000)

---

## 📊 DADOS DE TESTE - JOÃO CARLOS

### Estado Atual (Planilhas)

**Cliente**:
- ID: 1
- Nome: João Carlos Eduardo Figueiredo Basso
- CPF: XXX.XXX.XXX-XX
- Data Nascimento: 09/03/1974
- Idade Atual: 52 anos
- Tempo Contribuição: 13,88 anos

**Vínculos**: 13 registros (planilha Vinculos)

**Remunerações**: 32 registros (planilha Remuneracoes)

| Competência | Valor Original | Fator INPC (estimado) | Valor Corrigido |
|-------------|----------------|-----------------------|-----------------|
| 04/1998 | R$ 286 | ~5.2x | ~R$ 1.487 |
| 12/1999 | R$ 459 | ~3.4x | ~R$ 1.560 |
| 06/2004 | R$ 1.616 | ~2.1x | ~R$ 3.394 |
| ... | ... | ... | ... |

**Cálculo Esperado**:
- Média 80% maiores (SEM INPC): ~R$ 1.000 → todos cenários R$ 1.412 (mínimo)
- Média 80% maiores (COM INPC): ~R$ 3.500 → cenários R$ 2.100 / R$ 2.800 / R$ 3.500

---

## ⚠️ PONTOS DE ATENÇÃO

### 1. Detecção de Modo (Produção vs Desenvolvimento)

**Código**: modSimulacoes.AtualizarIndicesINPC(), linha 1149

```vb
If Dir(caminhoExe) <> "" Then
    ' PRODUÇÃO: Cliente sem Python
Else
    ' DESENVOLVIMENTO: Seu ambiente com Python
End If
```

**Validar**:
- [ ] No Parallels COM bin/atualizar_inpc.exe → usa .exe
- [ ] No Parallels SEM bin/ → usa Python (se instalado)
- [ ] Erro claro se nenhum disponível

---

### 2. Formato de Competência

**Crítico**: Competências devem estar no formato `MM/YYYY`

**Onde verificar**:
- Planilha "Remuneracoes", coluna D (4)
- Arquivo CSV importado: `extrato_..._remuneracoes.csv`

**Se errado** (ex: "1998-04"):
- Lookup falhará: `fatoresINPC.Exists("1998-04")` → False
- Valor retornado SEM correção
- Cálculos ficam errados

**Solução**: converter_extrato_inss.py já gera MM/YYYY

---

### 3. Separador CSV

**Python gera**: `;` (ponto e vírgula)
```csv
Competencia;Fator
01/1979;27.123456
```

**VBA lê**: `Split(linha, ";")`

**VBA converte número**: `CDbl(Replace("1.234567", ".", ","))`
- Ponto → vírgula (locale brasileiro)

**Validar**: teste manual do CSV gerado

---

### 4. Cache do Dictionary

**AplicarCorrecaoINPC()** usa Static:
- Primeira chamada: carrega CSV inteiro (~400 linhas)
- Demais chamadas: usa memória (instantâneo)
- Reseta se mudar de dia: `DateDiff("d", ultimaLeitura, Date) > 0`

**Implicação**: Se atualizar INPC durante o dia, pode precisar:
1. Fechar Excel
2. Reabrir
3. Calcular novamente

**OU** adicionar função "LimparCacheINPC()" se necessário

---

### 5. Timeout do Shell

**Código atual**: `Application.Wait (5 segundos)`

**Riscos**:
- BCB lento: pode não completar em 5s
- CSV não criado → erro falso

**Solução futura**: verificar processo em loop:
```vb
Do While Dir(arquivoSaida) = "" AND tentativas < 20
    Application.Wait (Now + TimeValue("00:00:01"))
    tentativas = tentativas + 1
Loop
```

---

## ✅ CHECKLIST DE CÓDIGO

### Arquivos Python

- [ ] `atualizar_inpc.py` - Sintaxe OK
- [ ] `atualizar_inpc.py` - Testa conexão BCB
- [ ] `atualizar_inpc.py` - Gera CSV corretamente
- [ ] `converter_extrato_inss.py` - Sintaxe OK
- [ ] `converter_extrato_inss.py` - Lê PDF sem erros
- [ ] `build_executaveis.py` - Gera .exe sem erros

### Módulos VBA

- [ ] `modSimulacoes.bas` - Compila sem erros
- [ ] `modSimulacoes.AtualizarIndicesINPC()` - Detecta .exe vs .py
- [ ] `modSimulacoes.AplicarCorrecaoINPC()` - Lê CSV corretamente
- [ ] `modSimulacoes.CalcularMediaSalarios()` - Aplica INPC
- [ ] `modDB.SetParametro()` - Salva em Config_Regras
- [ ] `modDB.GetParametro()` - Lê de Config_Regras

### Formulários VBA

- [ ] `frmPrincipal` - Abre sem erros
- [ ] `frmSimulacoes` - Botão "Atualizar INPC" visível
- [ ] `frmSimulacoes` - Botão "Calcular" funciona
- [ ] `frmCadastro` - Salva clientes corretamente

### Integração

- [ ] VBA → Python/exe: Shell funciona
- [ ] Python → CSV: arquivo gerado
- [ ] VBA → CSV: leitura correta
- [ ] Cache: Dictionary carrega 1x

---

## 🔐 SEGURANÇA E VALIDAÇÕES

### Validações Implementadas

| Local | Validação | Tratamento |
|-------|-----------|------------|
| AtualizarIndicesINPC | Verifica existência .exe/.py | MsgBox erro + Exit Function |
| AtualizarIndicesINPC | Verifica criação CSV | MsgBox erro + Exit Function |
| AplicarCorrecaoINPC | Verifica existência CSV | Retorna valor sem correção |
| AplicarCorrecaoINPC | Competência não encontrada | Retorna valor sem correção |
| CalcularMediaSalarios | Nenhum salário encontrado | Retorna 0 |

### Validações Pendentes

- [ ] Timeout inteligente no Shell (em vez de 5s fixo)
- [ ] Validar formato MM/YYYY antes de lookup
- [ ] Log de erros em arquivo texto
- [ ] Tratamento de permissões de arquivo

---

## 📈 PERFORMANCE

### Métricas Esperadas

| Operação | Tempo | Observações |
|----------|-------|-------------|
| Abrir Excel | < 3s | Depende do hardware |
| Atualizar INPC | 5-15s | Depende da internet |
| Carregar cache INPC | < 1s | Primeira vez após abrir Excel |
| Calcular simulação (1 cliente) | < 2s | 32 remunerações |
| Calcular simulação (100 clientes) | ⏳ Não testado | - |

### Otimizações Implementadas

- ✅ Static Dictionary (cache INPC)
- ✅ Leitura CSV 1x por dia
- ✅ Bubble sort limitado (apenas 32 itens)

### Otimizações Futuras

- [ ] Índice de clientes em memória
- [ ] Cache de simulações (evitar recálculo)
- [ ] Processamento assíncrono (VBA limitado)

---

## 🧪 MATRIZ DE TESTES

### Testes Unitários (Funções Isoladas)

| Função | Input | Output Esperado | Status |
|--------|-------|-----------------|--------|
| AplicarCorrecaoINPC | (286, "04/1998") | ~1.487 | ⏳ |
| AplicarCorrecaoINPC | (286, "99/9999") | 286 (sem correção) | ⏳ |
| CalcularMediaSalarios | ID=1 | ~3.500 | ⏳ |
| AtualizarIndicesINPC | - | True + CSV criado | ⏳ |

### Testes Integração (Fluxos Completos)

| Fluxo | Passos | Resultado Esperado | Status |
|-------|--------|-------------------|--------|
| Primeiro uso | Abrir Excel → Habilitar macros | Formulário principal | ⏳ |
| Atualizar INPC | Simulações → Atualizar INPC | MsgBox sucesso + CSV | ⏳ |
| Calcular simulação | Selecionar João → Calcular | 3 cenários com valores | ⏳ |
| Importar PDF | Importar → Selecionar PDF | Dados nas planilhas | ⏳ |

---

## 🚀 PRÓXIMOS PASSOS

1. **Gerar executáveis**:
   ```bash
   python build_executaveis.py
   ```

2. **Testar no Windows Parallels**:
   - Seguir CHECKLIST_ENVIO.md seção por seção

3. **Validar João Carlos**:
   - Cálculo SEM INPC: ~R$ 1.412
   - Cálculo COM INPC: ~R$ 2.100-3.500

4. **Preparar envio**:
   - Limpar pasta saida/
   - Comprimir em .zip
   - Anexar INSTALACAO_CLIENTE.md

5. **Enviar ao escritório**:
   - E-mail com instruções
   - Solicitar feedback em 48h

---

**Documento criado**: 13/01/2026  
**Versão do sistema**: 1.0.0  
**Status**: ✅ Mapeamento completo
