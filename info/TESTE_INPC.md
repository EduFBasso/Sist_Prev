# TESTE DE INTEGRAÇÃO INPC - INSTRUÇÕES

## ⚠️ AMBIENTE: Windows via Parallels

**IMPORTANTE**: Este sistema roda em Windows (Parallels) dentro do macOS:

- VBA executa no Windows (não no Mac host)
- Python deve estar instalado **no Windows**
- Caminhos de arquivo são Windows (pasta compartilhada)
- Comando: `python` (não `python3`)

---

## Implementações Concluídas ✅

### 1. Script Python: `atualizar_inpc.py`

- **Localização**: Pasta principal do projeto
- **Função**: Baixa série 188 (INPC) do Banco Central
- **Saída**: `saida/inpc_fatores.csv`
- **Base**: Dezembro/2025 = 1.0000

### 2. Função VBA: `AtualizarIndicesINPC()`

- **Módulo**: modSimulacoes
- **Execução**: Shell command (python3 no Mac, python no Windows)
- **Validação**: Verifica criação do arquivo CSV
- **Persistência**: Salva data em Config_Regras (Data_Atualizacao_INPC)

### 3. Função VBA: `AplicarCorrecaoINPC()`

- **Módulo**: modSimulacoes (private)
- **Cache**: Dictionary estático para performance
- **Lógica**: Valor × Fator → Valor Corrigido

### 4. Modificação: `CalcularMediaSalarios()`

- **Mudança**: Aplica correção INPC antes do cálculo
- **Processo**:
  1. Lê valor original da remuneração
  2. Lê competência (MM/YYYY)
  3. Chama AplicarCorrecaoINPC(valor, competencia)
  4. Usa valor corrigido no array de salários
  5. Calcula média dos 80% maiores corrigidos

---

## ROTEIRO DE TESTE

### ⚙️ PRÉ-REQUISITO: Python Instalado no Windows

Verificar instalação do Python no Windows (Parallels):

1. Abrir **Prompt de Comando** (cmd) no Windows
2. Executar: `python --version`
3. Deve mostrar: `Python 3.x.x`

Se não instalado:

- Baixar: https://www.python.org/downloads/windows/
- Instalar marcando "Add Python to PATH"
- Reiniciar cmd e testar novamente

### Teste 1: Executar Script Python no Windows

**Abrir Prompt de Comando (cmd)** no Windows:

```cmd
cd C:\Users\[seu_usuario]\Documents\Sist_Prev
python atualizar_inpc.py
```

_Nota: Ajustar caminho conforme configuração da pasta compartilhada Parallels_

**Resultado Esperado**:

- Conecta ao BCB
- Baixa ~400+ registros INPC
- Calcula fatores com base 12/2025
- Cria `saida/inpc_fatores.csv`
- Mostra exemplos:
  - 01/1995: ~6.5x (R$ 500 → ~R$ 3.250)
  - 12/1999: ~3.2x
  - 12/2020: ~1.3x
  - 12/2025: 1.0x (BASE)

### Teste 2: Chamar Função VBA no Excel

**Opção A - Janela Imediata (Ctrl+G)**:

```vb
?AtualizarIndicesINPC()
```

**Opção B - Criar Sub Temporária**:

```vb
Sub TestarINPC()
    Dim sucesso As Boolean
    sucesso = AtualizarIndicesINPC()

    If sucesso Then
        Debug.Print "Sucesso! Arquivo criado."
    Else
        Debug.Print "Falha na atualização."
    End If
End Sub
```

**Resultado Esperado**:

- Terminal Python abre e executa
- Aguarda 5 segundos
- Verifica arquivo CSV criado
- Salva data em Config_Regras
- MsgBox de sucesso

### Teste 3: Executar Simulação João Carlos

**Passo a Passo**:

1. Abrir frmSimulacoes
2. Selecionar "João Carlos Eduardo Figueiredo Basso" (ID=1)
3. Clicar "Calcular Simulações"

**ANTES (sem INPC)**:

```
Cenário Rápido:  R$ 1.412,00
Cenário Equilibrado: R$ 1.412,00
Cenário Máximo: R$ 1.412,00
```

Média calculada: ~R$ 1.000 (valores 1995-2004 sem correção)

**DEPOIS (com INPC)**:

```
Cenário Rápido:  R$ 2.100 - 2.500
Cenário Equilibrado: R$ 2.800 - 3.200
Cenário Máximo: R$ 3.500 - 4.200
```

Média calculada: ~R$ 3.500 (valores corrigidos para 12/2025)

### Teste 4: Verificar Arquivo CSV

**Abrir**: `saida/inpc_fatores.csv`

**Estrutura**:

```csv
Competencia;Fator
01/1979;27.123456
...
04/1998;5.234567
12/1999;3.456789
...
12/2025;1.000000
```

**Validações**:

- Fatores decrescentes (histórico → atual)
- Base 12/2025 = 1.000000
- Formato MM/YYYY consistente

### Teste 5: Verificar Config_Regras

**Planilha**: Config_Regras
**Nova Linha**: Data_Atualizacao_INPC | 13/01/2026 | Data da ultima... | 13/01/2026

---

## TROUBLESHOOTING

### ⚠️ Erro: Python não encontrado

**Sintoma**: VBA Shell mostra erro ou não executa
**Solução**:

1. Abrir cmd no Windows: `python --version`
2. Se não funcionar, instalar Python no Windows (não no Mac!)
3. Marcar opção "Add Python to PATH" durante instalação
4. Reiniciar Excel e testar novamente

### Erro: "Script atualizar_inpc.py não encontrado"

- Verificar arquivo na pasta principal (mesma do erp_prev.xlsm)
- No Windows, caminho deve ser acessível (pasta compartilhada Parallels)
- Testar manualmente: abrir cmd e navegar até a pasta

### Erro: "Arquivo inpc_fatores.csv não foi gerado"

- Executar Python manualmente no cmd (Windows)
- Verificar saída de erros no terminal
- Possível: falta biblioteca `requests` (não usada, usa urllib)
- Possível: problema de conectividade BCB

### Erro: Valores ainda R$ 1.412

- Verificar se CSV foi criado em `saida/inpc_fatores.csv`
- Verificar formato competência em Remuneracoes (deve ser MM/YYYY)
- Debug: adicionar `Debug.Print valorCorrigido` no CalcularMediaSalarios
- Exemplo João Carlos:
  - 04/1998: R$ 286 → ~R$ 1.500 (fator ~5.2x)
  - 12/1999: R$ 459 → ~R$ 1.585 (fator ~3.4x)

### CSV não é lido pelo VBA

- Verificar separador `;` (ponto e vírgula)
- Verificar formato de número: `1.234567` no CSV → `CDbl(Replace("1.234567", ".", ","))`
- Locale brasileiro usa vírgula como decimal

---

## PRÓXIMOS PASSOS

Após testes bem-sucedidos:

1. **Adicionar botão em frmSimulacoes**:

   - Nome: `cmdAtualizarINPC`
   - Caption: "Atualizar Índices INPC"
   - Evento: `Call AtualizarIndicesINPC()`

2. **Adicionar label de status**:

   - Nome: `lblDataINPC`
   - Caption: "Última atualização: [carregar do Config_Regras]"
   - UserForm_Initialize: `lblDataINPC.Caption = "Última atualização: " & GetParametro("Data_Atualizacao_INPC")`

3. **Testar cenário completo**:
   - PDF → converter_extrato_inss.py → CSV
   - Importar via VBA
   - Atualizar INPC
   - Simular benefícios
   - Validar valores realistas

---

## OBSERVAÇÕES TÉCNICAS

### Ambiente Windows/Parallels

- VBA executa no Windows (Parallels), não no macOS host
- Python deve estar instalado no Windows (comando: `python`)
- Pasta compartilhada: arquivos acessíveis por ambos os sistemas
- Application.PathSeparator detecta automaticamente o separador correto

### Performance

- `AplicarCorrecaoINPC()` usa Static Dictionary (cache)
- CSV lido 1x por dia (DateDiff validação)
- ~400 fatores em memória (~10KB)

### Manutenção

- Atualizar INPC mensalmente (automático via BCB API)
- Base 12/2025 fixa no código
- Para mudar base: editar `atualizar_inpc.py` linha 56

### Compatibilidade

- Windows: comando `python` (VBA executa em Parallels)
- CSV: separador `;` (padrão brasileiro)
- Encoding: UTF-8 no Python, ANSI no VBA
- Pasta compartilhada Parallels: sincronização automática entre Mac/Windows

Data do documento: 13/01/2026
