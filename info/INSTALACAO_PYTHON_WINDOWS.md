# INSTALAÇÃO PYTHON NO WINDOWS (PARALLELS)

## ⚠️ IMPORTANTE

Este sistema roda em **Windows via Parallels**. O Python deve estar instalado **no Windows**, não no macOS host.

---

## VERIFICAR SE JÁ ESTÁ INSTALADO

1. Abrir **Prompt de Comando** (cmd):

   - Pressionar `Win + R`
   - Digitar: `cmd`
   - Pressionar Enter

2. Executar:

   ```cmd
   python --version
   ```

3. **Se mostrar versão** (ex: `Python 3.11.5`):

   - ✅ Python já instalado!
   - Pular para seção "TESTAR SCRIPT"

4. **Se mostrar erro** ('python' não é reconhecido):
   - ⚠️ Python não instalado
   - Seguir passos de instalação abaixo

---

## INSTALAÇÃO DO PYTHON 3

### Passo 1: Download

1. Abrir navegador no Windows (Edge, Chrome, Firefox)
2. Acessar: https://www.python.org/downloads/
3. Clicar em **"Download Python 3.x.x"** (botão amarelo)
4. Salvar arquivo: `python-3.x.x-amd64.exe`

### Passo 2: Instalação

1. Executar arquivo baixado (duplo clique)
2. **⚠️ CRÍTICO**: Marcar checkbox **"Add Python to PATH"**
   ```
   [X] Add Python 3.x to PATH
   ```
3. Clicar em **"Install Now"**
4. Aguardar conclusão (2-3 minutos)
5. Clicar em **"Close"**

### Passo 3: Verificação

1. **FECHAR** todos os Prompt de Comando abertos
2. Abrir **NOVO** Prompt de Comando
3. Executar:
   ```cmd
   python --version
   ```
4. Deve mostrar: `Python 3.x.x`

---

## TESTAR SCRIPT INPC

Após instalação bem-sucedida:

1. Abrir Prompt de Comando (cmd)

2. Navegar para pasta do projeto:

   ```cmd
   cd C:\Users\[seu_usuario]\Documents\Sist_Prev
   ```

   _Nota: Substituir `[seu_usuario]` pelo nome real_

3. Executar script:

   ```cmd
   python atualizar_inpc.py
   ```

4. **Resultado esperado**:

   ```
   ============================================================
   ATUALIZAÇÃO DE ÍNDICES INPC - ERP_PREV
   ============================================================
   Data/Hora: 13/01/2026 14:30:45

   Conectando à API do Banco Central...
   ✓ 400+ registros obtidos com sucesso

   Calculando fatores com base em 12/2025...
   ✓ Fatores calculados para 400+ competências

   Gravando arquivo saida/inpc_fatores.csv...
   ✓ Arquivo salvo com 400+ registros

   --- Exemplos de fatores calculados ---
   Janeiro/1995: 6.543210x
   Dezembro/1999: 3.456789x
   Junho/2010: 1.876543x
   Dezembro/2020: 1.345678x
   Dezembro/2025 (BASE): 1.000000x

   ============================================================
   ✓ ATUALIZAÇÃO CONCLUÍDA COM SUCESSO!
   ============================================================

   Pressione ENTER para sair...
   ```

5. Verificar arquivo criado:
   ```cmd
   dir saida\inpc_fatores.csv
   ```

---

## TROUBLESHOOTING

### Erro: "python não é reconhecido como comando interno"

**Causa**: Python não foi adicionado ao PATH

**Solução 1 - Reinstalar**:

1. Desinstalar Python pelo Painel de Controle
2. Reinstalar marcando "Add Python to PATH"

**Solução 2 - Adicionar PATH manualmente**:

1. Localizar pasta Python (geralmente `C:\Users\[usuario]\AppData\Local\Programs\Python\Python3x`)
2. Variáveis de Ambiente → PATH → Adicionar caminho

### Erro: "No module named 'urllib'"

**Não deve ocorrer** - urllib é biblioteca padrão do Python

### Erro: "Cannot connect to BCB API"

**Causa**: Problema de rede/firewall

**Solução**:

1. Verificar conexão internet no Windows
2. Testar no navegador: https://api.bcb.gov.br/dados/serie/bcdata.sgs.188/dados?formato=json
3. Verificar firewall/antivírus bloqueando Python

### Script executa mas não cria CSV

**Verificar**:

1. Pasta `saida` existe? Se não, script cria automaticamente
2. Permissões de escrita na pasta?
3. Executar como administrador (botão direito → "Executar como administrador")

---

## EXECUTAR VIA VBA

Após Python instalado e testado manualmente:

1. Abrir Excel (erp_prev.xlsm)
2. Alt + F11 (Editor VBA)
3. Janela Imediata (Ctrl + G)
4. Executar:
   ```vb
   ?AtualizarIndicesINPC()
   ```
5. Deve abrir terminal, executar script e retornar `True`

---

## DICAS

### Caminho da Pasta Compartilhada

Parallels geralmente mapeia para:

- `C:\Users\[usuario]\Documents\Sist_Prev` (se configurado como compartilhada)
- Ou: verificar em Parallels → Configurações → Opções → Pastas Compartilhadas

### Testar Caminho no VBA

No Excel (Alt + F11, Janela Imediata):

```vb
?ThisWorkbook.Path
```

Mostra caminho completo onde Excel está rodando.

### Python 32 bits vs 64 bits

- Usar versão que corresponde ao Windows
- Geralmente: **64 bits** (mais comum)
- Verificar: Painel de Controle → Sistema → Tipo de sistema

---

Data: 13/01/2026
