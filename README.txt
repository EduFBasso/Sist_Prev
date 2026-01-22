═══════════════════════════════════════════════════════════════════════════════════════
   ███████╗██╗███████╗████████╗    ██████╗ ██████╗ ███████╗██╗   ██╗
   ██╔════╝██║██╔════╝╚══██╔══╝    ██╔══██╗██╔══██╗██╔════╝██║   ██║
   ███████╗██║███████╗   ██║       ██████╔╝██████╔╝█████╗  ██║   ██║
   ╚════██║██║╚════██║   ██║       ██╔═══╝ ██╔══██╗██╔══╝  ╚██╗ ██╔╝
   ███████║██║███████║   ██║       ██║     ██║  ██║███████╗ ╚████╔╝ 
   ╚══════╝╚═╝╚══════╝   ╚═╝       ╚═╝     ╚═╝  ╚═╝╚══════╝  ╚═══╝  
                                                                       
   Sistema de Planejamento Previdenciário - RGPS
   Versão 1.1 | Janeiro 2026
═══════════════════════════════════════════════════════════════════════════════════════

📁 ESTRUTURA DO PROJETO

Sist_Prev/
│
├── 📂 entrada/              ← COLOQUE OS PDFs CNIS AQUI
│   └── LEIA-ME.txt
│
├── 📂 saida/                ← CSVs GERADOS APARECEM AQUI
│   └── LEIA-ME.txt
│
├── 📂 info/                 ← DOCUMENTAÇÃO COMPLETA DO SISTEMA
│   ├── README.md
│   ├── MAPEAMENTO_FRMSIMULACOES.md
│   ├── MAPA_REGRAS_E_LAYOUT.md
│   ├── CORRECOES_IMPLEMENTADAS.md
│   ├── STATUS_APOS_CORRECOES.md
│   └── ... (outros manuais)
│
├── 📂 Formularios/          ← Código VBA dos formulários
│   ├── frmCadastro.bas
│   ├── frmSimulacoes.bas
│   ├── frmBusca.bas
│   └── ...
│
├── 📂 Modulos/              ← Código VBA das funções de cálculo
│   ├── modSimulacoes.bas
│   ├── modCadastro.bas
│   ├── modUtil.bas
│   └── ...
│
├── 📂 Planilhas/            ← Configurações e parâmetros
│   └── Config_Regras.bas
│
├── 🐍 converter_extrato_inss.py   ← EXTRATOR PYTHON (v1.1)
├── � atualizar_inpc.py           ← ATUALIZA ÍNDICE INPC (BCB)
├── 🐍 atualizar_selic.py          ← ATUALIZA ÍNDICE SELIC (BCB)
├── 🐍 sync_vba.py                 ← SINCRONIZA CÓDIGO VBA
├── �📊 erp_prev.xlsm                ← PLANILHA PRINCIPAL
└── 📖 COMO_USAR.txt                ← GUIA RÁPIDO

═══════════════════════════════════════════════════════════════════════════════════════

🚀 INÍCIO RÁPIDO

  1️⃣  PROCESSAR CNIS:
      • Coloque PDFs em "entrada/"
      • Execute (modo dev): python converter_extrato_inss.py
      • Execute (modo cliente/Windows): bin\converter_extrato_inss.exe
      • CSVs gerados em "saida/"

  2️⃣  IMPORTAR DADOS:
      • Abrir erp_prev.xlsm
      • Importar CSVs via VBA

  3️⃣  SIMULAR APOSENTADORIA:
      • frmBusca → Buscar cliente
      • frmCadastro → Ver dados
      • Botão "Simular" → frmSimulacoes
      • Botão "Calcular" → Ver resultados

═══════════════════════════════════════════════════════════════════════════════════════

✨ NOVIDADES VERSÃO 1.1 (12/01/2026)

  ✅ Função Nz() totalmente compatível com Access VBA
  ✅ Validação obrigatória de data de nascimento
  ✅ Cálculo preciso de tempo na reforma (sem estimativas)
  ✅ Análise automática da melhor regra de aposentadoria
  ✅ Estrutura organizada de pastas (entrada/saida/info)
  ✅ Processamento em lote de múltiplos PDFs
  ✅ Timestamps nos arquivos para evitar sobrescrever

═══════════════════════════════════════════════════════════════════════════════════════

📚 DOCUMENTAÇÃO

  Para informações detalhadas, consulte:
  
  • COMO_USAR.txt                          ← Guia completo de uso
  • info/README.md                         ← Visão geral do sistema
  • info/MAPEAMENTO_FRMSIMULACOES.md      ← Detalhes técnicos
  • info/STATUS_APOS_CORRECOES.md         ← Checklist de testes

═══════════════════════════════════════════════════════════════════════════════════════

💻 REQUISITOS

  Python:
    • Modo dev (rodar .py): Python 3.8+ e bibliotecas (ver requirements.txt)
    • Modo cliente (rodar .exe): NÃO precisa Python (usar bin\*.exe)
  
  Excel:
    • Microsoft Excel com suporte a macros VBA
    • Planilhas: Cadastro_Clientes, Vinculos, Config_Regras

═══════════════════════════════════════════════════════════════════════════════════════

🎯 FUNCIONALIDADES PRINCIPAIS

  ✓ Extração automática de dados do CNIS (PDF → CSV)
  ✓ Cadastro completo de clientes e vínculos
  ✓ Cálculo de tempo total de contribuição
  ✓ Conversão de tempo especial (fatores 1.4, 1.75, 2.33)
  ✓ 5 regras de aposentadoria implementadas:
    - Tempo de Contribuição (35H/30M)
    - Idade Mínima (65H/62M)
    - Regra de Pontos (105H/100M)
    - Pedágio 50%
    - Pedágio 100%
  ✓ Simulação de 3 cenários (15, 25, 40 anos)
  ✓ Identificação automática da melhor regra
  ✓ Projeção de valores de benefícios

═══════════════════════════════════════════════════════════════════════════════════════

🔧 SUPORTE E MANUTENÇÃO

  Bugs Conhecidos: Nenhum (v1.1)
  
  Melhorias Futuras:
    • Cálculo real de valores com remunerações históricas
    • Suporte a vínculos facultativos e contribuinte individual
    • Reajustes de valores por índices INPC/TR
    • Empacotamento Windows (.exe) via build_executaveis.py (padronizar entrega em bin/)

═══════════════════════════════════════════════════════════════════════════════════════

🪟 WINDOWS / ENTREGA AO CLIENTE (RECOMENDADO)

  Para rodar no Windows sem instalar Python:
    • Gere os executáveis com build_executaveis.py
    • Copie para a pasta bin/ ao lado do erp_prev.xlsm:
        - bin\atualizar_inpc.exe
        - bin\converter_extrato_inss.exe

  Observação:
    • O VBA procura primeiro o .exe em bin/. Se não existir, tenta rodar o .py (exige Python no Windows).

═══════════════════════════════════════════════════════════════════════════════════════

📧 CONTATO

  Para dúvidas, sugestões ou reportar problemas:
  • Consulte primeiro a pasta "info/" com a documentação
  • Verifique o arquivo *_validacao.txt para erros de extração

═══════════════════════════════════════════════════════════════════════════════════════

© 2026 - Sist_Prev | Sistema de Planejamento Previdenciário
Desenvolvido para advogados previdenciários

═══════════════════════════════════════════════════════════════════════════════════════
