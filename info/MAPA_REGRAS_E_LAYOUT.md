# Mapa de Regras e Layout (referência para refatoração)

## Objetivo

Este documento separa o que é **regra por vínculo** vs **regra por simulação** e descreve como o sistema já está organizado hoje, para permitir refatorar layout (telas menores) sem quebrar o motor.

> Observação importante: este sistema é técnico. A definição do que “a lei manda” é jurídica.
> O software deve ser **parametrizável** para acomodar divergências (INPC vs SELIC, fatores especiais, etc.).

---

## 1) Onde cada regra deve morar

### 1.1 Regras por VÍNCULO (características do período)

Essas regras pertencem ao registro de vínculo e afetam o cálculo de tempo:

- **Especial (Sim/Não)** + **Grau (15/20/25)**
  - Entrada: `Vinculos` col 6 (Especial) e col 7 (Grau)
  - Cálculo: `modSimulacoes.CalcularTempoEspecial()`
  - Fatores: parametrizados via `Config_Regras` (ex.: `Conversao_Especial_25_H`, `Conversao_Especial_25_M`, etc.)

- Flags do vínculo (Rural/Militar/Exterior/Atraso/Complementar)
  - Entrada: checkboxes em `frmVinculos` e colunas 11–16 da planilha `Vinculos`
  - Uso: atualmente como indicadores; podem virar regras futuras.

Conclusão prática: **“Especial” deve ser marcado no vínculo** (é uma característica daquele período).

### 1.2 Regras por SIMULAÇÃO (regras de aposentadoria)

Essas regras comparam tempo, idade, carência e transições:

- Tempo de contribuição, Idade, Pontos, Pedágio 50%, Pedágio 100%
  - Motor: `modSimulacoes.RegraTempoContribuicao`, `RegraIdade`, `RegraPontos`, `RegraPedagio50`, `RegraPedagio100`
  - Orquestração: `modSimulacoes.AnalisarMelhorRegra()`
  - UI: `frmSimulacoes` (apenas exibição e gatilho)

Conclusão prática: **as regras comuns ficam no motor (`modSimulacoes`) e o `frmSimulacoes` só apresenta**.

---

## 2) Checkbox no frmCadastro: remover ou manter?

No `frmCadastro`, os checkboxes `chkPossuiEspecial`, `chkPossuiRural`, etc. **não são “tipos” para o usuário escolher**.
Eles são **indicadores automáticos** do conjunto de vínculos do cliente.

Como funciona hoje:

- Ao salvar/excluir vínculo: `modVinculos.AtualizarIndicadoresCliente(ID_Cliente)` atualiza colunas 22–28 na planilha `Cadastro_Clientes`.
- Ao carregar cliente: `modCadastro.CarregarCliente()` preenche os checkboxes do `frmCadastro` com base nessas colunas.

Recomendação de layout (tela menor):

- Manter como **somente leitura** (desabilitados) OU substituir por labels/ícones.
- Se precisar ganhar espaço vertical, este é um bloco bom para virar “Resumo” colapsável.

---

## 3) “Tipo de Segurado” (ComboBox) vs múltiplos tipos

- `frmCadastro.cboTipoSegurado` é **1 de N** (single-select) → ComboBox é o controle correto.
- Se no futuro houver **multi-select real** (ex.: filtrar clientes por vários tipos), prefira:
  - `ListBox.MultiSelect` (economiza espaço e suporta muitos valores)
  - ou um popup “Selecionar tipos…” em vez de 8–12 checkboxes no formulário.

---

## 4) Índice de correção monetária (INPC vs SELIC)

Hoje o sistema aplica **INPC** na correção dos salários históricos (remunerações) no cálculo de média:

- `modSimulacoes.CalcularMediaSalarios()` → aplica correção por competência
- INPC é atualizado via Python (`atualizar_inpc.py` / `bin/atualizar_inpc.exe`) e gravado em `saida/inpc_fatores.csv`

Para lidar com divergência (advogado vs advogada), o caminho de engenharia é:

- **tornar o índice selecionável por parâmetro** e isolar em um ponto único.

Parâmetro sugerido (Config_Regras):

- `Indice_Correcao_Remuneracoes` = `INPC` (padrão) ou `SELIC` (futuro)

Status atual no VBA:

- `INPC`: implementado
- `SELIC`: placeholder (não implementado; valores ficam sem correção se selecionado)

---

## 5) Checklist para fechar com o jurídico

Sugestão de perguntas objetivas para o advogado decidir (para virar parâmetros/regra no sistema):

1. Tempo especial

- Em quais situações o cliente **pode converter** especial→comum?
- Quando o correto é **aposentadoria especial** (sem conversão) vs **conversão** para regra comum?

2. Carência

- Contar carência por competências (mensal) ou por período?

3. Correção monetária

- INPC ou SELIC, e a partir de qual marco temporal?
- É correção dos salários para média ou correção de atrasados? (são coisas diferentes)

4. Sexo

- Usar sexo do cadastro como fonte única? E como tratar inconsistência ("Masculino" vs "M")?

---

## 6) Para refatorar layout sem risco

- Não mexer no motor (`modSimulacoes`, `modVinculos`) durante o redesenho.
- No `frmCadastro`:
  - Tratar checkboxes “Possui…” como indicadores (read-only)
  - Priorizar campos essenciais visíveis sem scroll
- No `frmSimulacoes`:
  - Manter botões e resultados; regras ficam no módulo
