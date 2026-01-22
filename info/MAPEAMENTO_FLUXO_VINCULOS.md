# 🗺️ MAPEAMENTO COMPLETO DO SISTEMA - Fluxo de Vínculos

**Data:** 21 de janeiro de 2026  
**Sistema:** Sist_Prev v1.1  
**Foco:** Arquitetura e Fluxo de Vínculos

---

## 📊 FLUXO PRINCIPAL DO SISTEMA

```
┌─────────────────┐
│  frmPrincipal   │ Menu principal do sistema
└────────┬────────┘
         │
         ├─> 🔍 BUSCAR CLIENTE
         │   ┌──────────────┐
         │   │  frmBusca    │ Busca cliente por nome/CPF/NIT
         │   └──────┬───────┘
         │          │ (Duplo clique na linha)
         │          ▼
         ├─> 📋 CADASTRO DO CLIENTE
         │   ┌──────────────────┐
         │   │  frmCadastro     │ Dados pessoais + lista de vínculos
         │   └──────┬───────┬───┘
         │          │       │
         │          │       ├─> [Adicionar Vínculo]
         │          │       │   ┌────────────────┐
         │          │       └─> │  frmVinculos   │ ← AQUI!
         │          │           └────────────────┘
         │          │
         │          ├─> [Importar CNIS] ─> Python ─> CSVs ─> ImportarVinculosDeCSV()
         │          │
         │          └─> [Simular]
         │                  │
         │                  ▼
         └─> 📊 SIMULAÇÃO
             ┌──────────────────┐
             │  frmSimulacoes   │ Cálculos de aposentadoria
             └──────────────────┘
                     │
                     │ LEITURA DOS DADOS
                     ▼
             ┌──────────────────┐
             │ Planilha Vinculos│ Armazena todos os vínculos
             └──────────────────┘
                     │
                     │ PROCESSAMENTO
                     ▼
             ┌──────────────────┐
             │  modSimulacoes   │ Motor de cálculo
             └──────────────────┘
```

---

## 🏗️ ARQUITETURA DE VÍNCULOS

### Camadas do Sistema:

```
┌──────────────────────────────────────────────────────────────┐
│                    CAMADA DE APRESENTAÇÃO                     │
├──────────────────────────────────────────────────────────────┤
│  frmVinculos.bas                                              │
│  • Interface visual para adicionar/editar vínculos            │
│  • Validações de UI (habilitar/desabilitar cboGrau)          │
│  • Coleta dados do formulário → Collection                   │
└──────────────────────┬───────────────────────────────────────┘
                       │ ColetarDados()
                       ▼
┌──────────────────────────────────────────────────────────────┐
│                    CAMADA DE LÓGICA                          │
├──────────────────────────────────────────────────────────────┤
│  modVinculos.bas                                              │
│  • SalvarVinculo(dados)         - Grava/atualiza vínculo     │
│  • ExcluirVinculo(ID)           - Remove vínculo             │
│  • CarregarVinculo(ID)          - Carrega dados para edição  │
│  • CarregarVinculosCliente(ID)  - Lista vínculos do cliente  │
│  • AtualizarIndicadoresCliente  - Atualiza flags (Rural,etc) │
└──────────────────────┬───────────────────────────────────────┘
                       │ Escreve/Lê
                       ▼
┌──────────────────────────────────────────────────────────────┐
│                    CAMADA DE DADOS                           │
├──────────────────────────────────────────────────────────────┤
│  Planilha "Vinculos" (Excel)                                  │
│  • Estrutura: 18 colunas                                      │
│  • Col 1:  ID_Vinculo                                         │
│  • Col 2:  ID_Cliente                                         │
│  • Col 3:  Data_Inicio                                        │
│  • Col 4:  Data_Fim                                           │
│  • Col 5:  Tipo (Empregado, CI, etc)                         │
│  • Col 6:  Especial (Sim/Não)                                │
│  • Col 7:  Grau (15/20/25)                                    │
│  • Col 8:  Salario                                            │
│  • Col 9:  Observacoes                                        │
│  • Col 10: Empresa (importado do CNIS)                        │
│  • Col 11: Rural (Boolean)                                    │
│  • Col 12: Militar (Boolean)                                  │
│  • Col 13: Exterior (Boolean)                                 │
│  • Col 14: Concomitante (Boolean)                             │
│  • Col 15: Atraso (Boolean)                                   │
│  • Col 16: Complementar (Boolean)                             │
│  • Col 17: Seq (CNIS)                                         │
│  • Col 18: Codigo_Emp (CNIS)                                  │
└──────────────────────┬───────────────────────────────────────┘
                       │ Leitura para cálculo
                       ▼
┌──────────────────────────────────────────────────────────────┐
│                    CAMADA DE CÁLCULO                         │
├──────────────────────────────────────────────────────────────┤
│  modSimulacoes.bas                                            │
│  • CalcularTempo(ID_Cliente)        - Soma vínculos           │
│  • CalcularTempoEspecial(ID_Cliente)- Converte especial       │
│  • AnalisarMelhorRegra(ID_Cliente)  - Escolhe aposentadoria   │
└──────────────────────────────────────────────────────────────┘
```

---

## 🎨 frmVinculos - DETALHAMENTO COMPLETO

### Controles Visuais:

#### 1. **Identificação (Ocultos)**
```vb
txtIDCliente  ' ID do cliente (hidden)
txtIDVinculo  ' ID do vínculo (hidden, 0 = novo)
```

#### 2. **Dados Básicos do Vínculo**
```vb
cboTipo       ' Tipo: Empregado, CI, Avulso, Doméstico, etc
txtInicio     ' Data início (dd/mm/yyyy)
txtFim        ' Data fim (dd/mm/yyyy ou vazio se em andamento)
```

#### 3. **Frame: fraEspecial - Atividade Especial**
```vb
cboEspecial   ' "Sim" ou "Não"
cboGrau       ' "15", "20" ou "25" (habilitado apenas se Especial=Sim)
```

**Lógica:** 
- `cboEspecial_Change()`: Se "Sim" → habilita cboGrau
- Se "Não" → desabilita cboGrau e limpa valor

#### 4. **Frame: fraVinculo - Características do Período**
```vb
chkRural         ' Trabalho rural
chkMilitar       ' Serviço militar
chkExterior      ' Trabalho no exterior
chkConcomitante  ' Período concomitante (sobreposição)
chkAtraso        ' Atraso no pagamento
chkComplementar  ' Complementação (tempo especial)
```

#### 5. **Informações Adicionais**
```vb
txtSalario   ' Último salário (informativo, sem função)
txtObs       ' Observações livres do advogado
```

#### 6. **Botões de Ação**
```vb
cmdNovo      ' Limpa formulário para novo vínculo
cmdSalvar    ' Salva (insert ou update)
cmdExcluir   ' Remove vínculo
cmdVoltar    ' Fecha formulário
```

---

## 🔄 FLUXO DE DADOS - Adicionar/Editar Vínculo

### 📥 Cenário 1: ADICIONAR NOVO VÍNCULO

```
frmCadastro
    │
    └─> [Botão: Adicionar Vínculo]
            │
            ▼
        frmVinculos.Show
            │
            ├─> UserForm_Initialize()
            │   └─> Popula combos (Tipo, Especial, Grau)
            │       Desabilita cboGrau (padrão: Especial=Não)
            │
            ├─> txtIDCliente = ID do cliente (hidden)
            │   txtIDVinculo = "" (novo)
            │
            ├─> Usuário preenche:
            │   • Tipo, Datas, Especial?, Grau?
            │   • Checkboxes conforme características
            │   • Observações
            │
            └─> [cmdSalvar_Click]
                    │
                    ├─> dados = ColetarDados()
                    │   └─> Collection com todos os campos
                    │
                    ├─> modVinculos.SalvarVinculo(dados)
                    │   └─> Grava na planilha Vinculos
                    │       ProximoID = última linha + 1
                    │
                    ├─> modVinculos.AtualizarIndicadoresCliente(ID)
                    │   └─> Atualiza flags no Cadastro_Clientes
                    │       (chkPossuiEspecial, chkPossuiRural, etc)
                    │
                    ├─> MsgBox "Salvo com sucesso!"
                    │
                    └─> Recarrega lista de vínculos no frmCadastro
```

### 📝 Cenário 2: EDITAR VÍNCULO EXISTENTE

```
frmCadastro
    │
    └─> [Duplo clique na linha de vínculo]
            │
            ▼
        modVinculos.CarregarVinculo(ID_Vinculo)
            │
            └─> Busca dados na planilha Vinculos
                Preenche todos os controles do frmVinculos
                    │
                    ├─> txtIDVinculo = ID (não é 0)
                    ├─> Datas, Tipo, Especial, Grau
                    ├─> Checkboxes
                    └─> Observações
                        │
                        └─> Usuário edita e clica [Salvar]
                                │
                                └─> Atualiza linha existente na planilha
```

### 🗑️ Cenário 3: EXCLUIR VÍNCULO

```
frmVinculos (com vínculo carregado)
    │
    └─> [cmdExcluir_Click]
            │
            ├─> Valida: txtIDVinculo não vazio?
            │
            ├─> modVinculos.ExcluirVinculo(ID)
            │   └─> Remove linha da planilha Vinculos
            │
            ├─> MsgBox "Excluído com sucesso!"
            │
            └─> Recarrega lista no frmCadastro
```

---

## 📋 IMPORTAÇÃO AUTOMÁTICA - CNIS

```
frmCadastro
    │
    └─> [Botão: Importar CNIS]
            │
            ├─> Usuário seleciona PDF do CNIS
            │
            ├─> Python: converter_extrato_inss.py
            │   └─> Extrai dados do PDF
            │       • Cliente (NIT, CPF, Nome, Nascimento)
            │       • Vínculos (Empresa, Datas)
            │       • Remunerações por competência
            │       └─> Gera CSVs em /saida/
            │
            ├─> modImportacao.ImportarVinculosDeCSV()
            │   └─> Lê CSV de vínculos
            │       Para cada linha:
            │           • Cria Collection com dados
            │           • Tipo = "Empregado" (padrão)
            │           • Especial = "Não" (padrão)
            │           • Checkboxes = False (padrão)
            │           • Chama SalvarVinculo(dados)
            │
            └─> Vínculos aparecem na lista do frmCadastro
                    │
                    └─> Advogado pode editar cada um:
                        • Marcar Especial = Sim
                        • Definir Grau (15/20/25)
                        • Marcar checkboxes (Rural, etc)
```

---

## 🎯 ONDE REFATORAR VÍNCULOS?

### Opção Recomendada: **frmVinculos PRIMEIRO**

#### ✅ Melhorias Necessárias:

1. **Validações de Entrada**
   ```vb
   ' Ao salvar, validar:
   - Data Início obrigatória
   - Data Início < Data Fim (se Fim preenchido)
   - Se Especial=Sim, Grau obrigatório
   - Tipo de vínculo obrigatório
   ```

2. **Nova Funcionalidade: Finalidade do Tempo Especial**
   ```vb
   ' Adicionar Frame: fraFinalidadeEspecial
   optConverter     ' Converter para tempo comum (padrão atual)
   optAposentarEsp  ' Usar para aposentadoria especial pura
   
   ' Nova coluna na planilha: Finalidade_Especial
   ' Valores: "CONVERTER" ou "APOSENTAR"
   ```

3. **Lógica no modSimulacoes**
   ```vb
   ' CalcularTempoEspecial() precisa distinguir:
   Function CalcularTempoEspecial(ID_Cliente, Finalidade)
       ' Se Finalidade = "CONVERTER":
       '   Aplica fatores (2.33, 1.75, 1.40)
       '   Retorna tempo convertido
       ' 
       ' Se Finalidade = "APOSENTAR":
       '   Retorna tempo puro (sem conversão)
       '   Será usado na regra de Aposentadoria Especial
   ```

4. **Integração com frmSimulacoes**
   ```vb
   ' AnalisarMelhorRegra() precisa verificar:
   ' - Se cliente tem vínculos com Finalidade="APOSENTAR"
   ' - Se sim, considerar RegraAposentadoriaEspecial()
   ' - Comparar com outras regras
   ```

---

## 📐 ESTRUTURA ATUAL DA PLANILHA VINCULOS

| Col | Nome | Tipo | Origem | Uso |
|-----|------|------|--------|-----|
| 1 | ID_Vinculo | Int | Auto | PK |
| 2 | ID_Cliente | Int | FK | Link cliente |
| 3 | Data_Inicio | Date | Manual/CNIS | Cálculo tempo |
| 4 | Data_Fim | Date | Manual/CNIS | Cálculo tempo |
| 5 | Tipo | String | Manual | Classificação |
| 6 | Especial | String | Manual | Sim/Não |
| 7 | Grau | String | Manual | 15/20/25 |
| 8 | Salario | Double | Manual | Informativo |
| 9 | Observacoes | String | Manual | Anotações |
| 10 | Empresa | String | CNIS | Referência |
| 11 | Rural | Boolean | Manual | Flag |
| 12 | Militar | Boolean | Manual | Flag |
| 13 | Exterior | Boolean | Manual | Flag |
| 14 | Concomitante | Boolean | Manual | Flag |
| 15 | Atraso | Boolean | Manual | Flag |
| 16 | Complementar | Boolean | Manual | Flag |
| 17 | Seq | String | CNIS | Referência |
| 18 | Codigo_Emp | String | CNIS | Referência |

### 🆕 Nova Coluna Proposta:

| Col | Nome | Tipo | Valores | Descrição |
|-----|------|------|---------|-----------|
| 19 | Finalidade_Especial | String | "CONVERTER" / "APOSENTAR" | Define como usar o tempo especial |

---

## 🚀 PRÓXIMOS PASSOS RECOMENDADOS

### Fase 1: Refatorar frmVinculos ⭐
1. Adicionar OptionButtons para Finalidade (Converter vs Aposentar)
2. Adicionar validações completas
3. Atualizar ColetarDados() para incluir Finalidade
4. Atualizar modVinculos.SalvarVinculo() para nova coluna

### Fase 2: Refatorar modSimulacoes
1. Separar CalcularTempoEspecial em duas funções
2. Implementar RegraAposentadoriaEspecial()
3. Integrar com AnalisarMelhorRegra()

### Fase 3: Refatorar frmSimulacoes
1. Completar implementação do optEspecial
2. Exibir corretamente quando selecionado
3. Calcular benefício específico

### Fase 4: Testes
1. Testar todos os fluxos
2. Validar cálculos
3. Comparar com casos reais

---

## 📞 DECISÕES PENDENTES (com parceiro jurídico)

1. **Finalidade do Especial:**
   - [ ] Cliente pode escolher entre converter e aposentar especial?
   - [ ] Sistema escolhe automaticamente a melhor opção?
   - [ ] Sempre converter (regra atual)?

2. **Validações de PPP:**
   - [ ] Sistema deve validar PPP (Perfil Profissiográfico)?
   - [ ] Apenas alerta que precisa de PPP?
   - [ ] Ignora (responsabilidade do advogado)?

3. **Cálculo de Benefício Especial:**
   - [ ] Usa mesma fórmula das regras comuns?
   - [ ] Fórmula diferente?
   - [ ] Apenas informativo (não calcula valor)?

---

**Documento criado em:** 21/01/2026  
**Última atualização:** 21/01/2026  
**Versão:** 1.0
