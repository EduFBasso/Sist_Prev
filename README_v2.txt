═══════════════════════════════════════════════════════════════════════════════════════
   ███████╗██╗███████╗████████╗    ██████╗ ██████╗ ███████╗██╗   ██╗
   ██╔════╝██║██╔════╝╚══██╔══╝    ██╔══██╗██╔══██╗██╔════╝██║   ██║
   ███████╗██║███████╗   ██║       ██████╔╝██████╔╝█████╗  ██║   ██║
   ╚════██║██║╚════██║   ██║       ██╔═══╝ ██╔══██╗██╔══╝  ╚██╗ ██╔╝
   ███████║██║███████║   ██║       ██║     ██║  ██║███████╗ ╚████╔╝ 
   ╚══════╝╚═╝╚══════╝   ╚═╝       ╚═╝     ╚═╝  ╚═╝╚══════╝  ╚═══╝  
                                                                       
   Sistema de Planejamento Previdenciário - RGPS
   Versão 2.0 | Janeiro 2026
═══════════════════════════════════════════════════════════════════════════════════════

📁 ESTRUTURA DO PROJETO

Sist_Prev/
│
├── 📂 cnis/                 ← PDFs CNIS (teste e exemplos)
│   └── CNIS_JOAO_CARLOS.pdf
│
├── 📂 saida/                ← CSVs GERADOS APARECEM AQUI
│   └── (5 arquivos por PDF processado)
│
├── 📂 extrator/             ← MÓDULO DE EXTRAÇÃO MODULAR (v2.0) 🆕
│   ├── tipos/               ← Parsers por tipo de vínculo
│   │   ├── coordenador_remuneracoes.py  ← Orquestrador principal
│   │   ├── clt.py           ← Extrator CLT (3 campos)
│   │   ├── facultativo.py   ← Extrator Facultativo (5 campos)
│   │   ├── detector.py      ← Detecção automática de tipos
│   │   └── config_vinculos.py
│   └── __init__.py
│
├── 📂 tests/                ← TESTES AUTOMATIZADOS 🆕
│   ├── test_extrator.py     ← 12 testes (100% passando)
│   └── README.md            ← Guia de testes
│
├── 📂 info/                 ← DOCUMENTAÇÃO COMPLETA
│   ├── README.md
│   ├── ARCHITECTURE.md      ← Arquitetura do extrator 🆕
│   ├── MAPEAMENTO_CODIGO.md
│   └── ... (outros manuais)
│
├── 📂 Formularios/          ← Código VBA dos formulários
├── 📂 Modulos/              ← Código VBA das funções de cálculo
├── 📂 Planilhas/            ← Configurações e parâmetros
│
├── 🐍 converter_extrato_inss.py   ← EXTRATOR PYTHON (v2.0) ⭐
├── 🐍 atualizar_inpc.py           ← ATUALIZA ÍNDICE INPC (BCB)
├── 🐍 atualizar_selic.py          ← ATUALIZA ÍNDICE SELIC (BCB)
├── 🐍 sync_vba.py                 ← SINCRONIZA CÓDIGO VBA
├── 📊 erp_prev.xlsm               ← PLANILHA PRINCIPAL
└── 📖 CHANGELOG.txt               ← HISTÓRICO DE VERSÕES

═══════════════════════════════════════════════════════════════════════════════════════

🚀 INÍCIO RÁPIDO

  1️⃣  PROCESSAR CNIS:
      
      # Usando Python (modo desenvolvimento)
      python converter_extrato_inss.py cnis/CNIS_EXEMPLO.pdf saida/resultado.csv
      
      # Usando executável (modo produção)
      bin\converter_extrato_inss.exe cnis\CNIS_EXEMPLO.pdf saida\resultado.csv
      
      # Ver ajuda
      python converter_extrato_inss.py --help

  2️⃣  ARQUIVOS GERADOS (5 CSVs):
      
      ✅ resultado.csv                      → Tabelas brutas do PDF
      ✅ resultado_dados_cliente.csv        → Nome, NIT, CPF do filiado
      ✅ resultado_vinculos_brutos.csv      → Blocos de vínculos (texto)
      ✅ resultado_vinculos_estruturado.csv → 10 vínculos CLT parseados
      ✅ resultado_remuneracoes.csv         → 178 remunerações (CLT + Facultativo)

  3️⃣  RODAR TESTES:
      
      pytest tests/ -v                      → Executa 12 testes
      pytest tests/ --cov=converter_extrato_inss --cov=extrator  → Com cobertura

═══════════════════════════════════════════════════════════════════════════════════════

✨ NOVIDADES VERSÃO 2.0 (26/01/2026)

  🎯 EXTRAÇÃO COMPLETA (100%)
    ✅ 178/178 remunerações extraídas (163 CLT + 15 Facultativo)
    ✅ Todas 13 sequências de vínculos capturadas
    ✅ Valores soltos entre páginas (Seq 2, 8, 11)
    ✅ Blocos cortados no fim da página (Seq 7, 10, 11)

  🏗️ ARQUITETURA MODULAR
    ✅ Coordenador com delegação por tipo de vínculo
    ✅ Parser CLT (3 campos: Competência | Remuneração | Indicadores)
    ✅ Parser Facultativo (5 campos: + Data Pgto + Salário)
    ✅ Detecção automática de tipos (CLT vs Facultativo)
    ✅ Código reduzido: 911 → 587 linhas (-36%)

  🧪 TESTES AUTOMATIZADOS
    ✅ 12 testes implementados (100% passando)
    ✅ Cobertura de 48% do código
    ✅ Validação automática de 178 remunerações
    ✅ CI/CD ready (pytest + pytest-cov)

  💅 UX MELHORADA
    ✅ Mensagens com emojis (📄 ✅ ❌)
    ✅ Opção --help mostra documentação completa
    ✅ Validação de arquivo não encontrado
    ✅ Auto-criação de diretórios de saída

  📚 DOCUMENTAÇÃO COMPLETA
    ✅ ARCHITECTURE.md (fluxo de extração detalhado)
    ✅ tests/README.md (guia de testes)
    ✅ CHANGELOG.txt (histórico de mudanças)
    ✅ README.txt atualizado (este arquivo)

═══════════════════════════════════════════════════════════════════════════════════════

📊 TIPOS DE VÍNCULOS SUPORTADOS

  1️⃣  CLT (Vínculos com Empregador)
      • Detectado por: "Matrícula do Tipo Filiado" + "Código Emp." (CNPJ)
      • Seção: "Remunerações"
      • Formato: Competência | Remuneração | Indicadores (3 campos)
      • Até 3 competências por linha
      • Exemplo: 01/1995 | 286,25 | 13º SALÁRIO

  2️⃣  FACULTATIVO (Contribuinte Facultativo)
      • Detectado por: "Origem do Vínculo" + NIT + "RECOLHIMENTO"
      • Seção: "Contribuições"
      • Formato: Competência | Data Pgto | Contribuição | Salário | Indicadores (5 campos)
      • Até 2 competências por linha
      • Captura: Competência, Contribuição, Indicadores
      • Ignora: Data Pgto, Salário (presentes no PDF, não relevantes para cálculo)
      • Exemplo: 01/2020 | 15/02/2020 | 220,00 | 2.000,00 | PREC-FACULTCONC

═══════════════════════════════════════════════════════════════════════════════════════

💻 REQUISITOS

  Python (modo desenvolvimento):
    • Python 3.8+
    • Bibliotecas: pdfplumber, pytest (ver requirements.txt)
    • Instalar: pip install -r requirements.txt
  
  Executável (modo produção):
    • NÃO precisa Python instalado
    • Usar: bin\converter_extrato_inss.exe (Windows)
  
  Excel:
    • Microsoft Excel com suporte a macros VBA
    • Planilhas: Cadastro_Clientes, Vinculos, Config_Regras

═══════════════════════════════════════════════════════════════════════════════════════

🎯 FUNCIONALIDADES PRINCIPAIS

  EXTRAÇÃO DE DADOS
    ✓ Extração automática de CNIS (PDF → CSV)
    ✓ Suporte CLT e Facultativo
    ✓ Detecção automática de tipos
    ✓ Captura valores entre páginas
    ✓ 100% de precisão (178/178 remunerações)

  CÁLCULOS PREVIDENCIÁRIOS
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

═══════════════════════════════════════════════════════════════════════════════════════

🔧 DESENVOLVIMENTO

  RODAR TESTES:
    pytest tests/ -v                        # 12 testes
    pytest tests/ --cov=converter_extrato_inss --cov=extrator  # Com cobertura

  ESTRUTURA DE TESTES:
    ✓ test_extracao_total_178_remuneracoes  → Valida total exato
    ✓ test_distribuicao_por_seq             → Valida Seq 1-13
    ✓ test_vinculos_clt_tem_cnpj            → CNPJ válido
    ✓ test_vinculos_facultativo             → 15 registros (Seq 11-13)
    ✓ test_formato_competencia              → MM/AAAA
    ✓ test_valores_remuneracao_validos      → Numéricos positivos
    ✓ test_gerar_csv_remuneracoes           → 178 linhas CSV
    ✓ test_gerar_csv_dados_cliente          → 1 linha CSV
    ✓ test_gerar_csv_vinculos_estruturado   → 10 vínculos CLT
    ✓ test_pdf_inexistente                  → FileNotFoundError
    ✓ test_ordenacao_por_seq                → Seq crescente
    ✓ test_pipeline_completo                → 5 CSVs gerados

  COBERTURA DE CÓDIGO:
    • converter_extrato_inss.py: 71%
    • coordenador_remuneracoes.py: 88%
    • config_vinculos.py: 88%
    • tipos/clt.py: 68%
    • tipos/facultativo.py: 82%
    • TOTAL: 48%

═══════════════════════════════════════════════════════════════════════════════════════

🪟 ENTREGA AO CLIENTE (WINDOWS)

  Para rodar no Windows sem instalar Python:
    
    1. Gere os executáveis:
       python build_executaveis.py
    
    2. Arquivos gerados em bin/:
       • bin\converter_extrato_inss.exe
       • bin\atualizar_inpc.exe
       • bin\atualizar_selic.exe
    
    3. Distribua a pasta completa:
       Sist_Prev\
       ├── bin\*.exe
       ├── erp_prev.xlsm
       ├── cnis\      (exemplos)
       └── saida\     (vazio)

═══════════════════════════════════════════════════════════════════════════════════════

🚀 PRÓXIMOS PASSOS (ROADMAP)

  FASE 3: MOTOR DE CÁLCULO VIA PLANILHA
    • Exportar dados estruturados para Excel
    • Adaptar Config_Regras.bas para Python/Excel
    • Implementar cálculos com transparência (mostrar "por quê")
    • Validar contra sistema VBA (prova de conceito)
    • Documentar divergências e justificativas

  MELHORIAS FUTURAS:
    • Suporte a outros tipos de vínculos (Rural, Autônomo, CI, MEI)
    • Cálculo real de valores com remunerações históricas
    • Reajustes por índices INPC/TR/SELIC
    • Tratamento de períodos concomitantes (média de salários)
    • Interface gráfica (GUI) para arrastar PDF

═══════════════════════════════════════════════════════════════════════════════════════

📧 SUPORTE

  Para dúvidas, sugestões ou reportar problemas:
    • Consulte primeiro a pasta "info/" com a documentação completa
    • Verifique CHANGELOG.txt para mudanças recentes
    • Execute os testes: pytest tests/ -v
    • Veja logs em arquivos *_validacao.txt

  DOCUMENTAÇÃO ÚTIL:
    • README.txt                         → Este arquivo (visão geral)
    • CHANGELOG.txt                      → Histórico de mudanças
    • info/ARCHITECTURE.md               → Arquitetura do extrator
    • info/MAPEAMENTO_CODIGO.md          → Mapa de funções VBA
    • tests/README.md                    → Guia de testes

═══════════════════════════════════════════════════════════════════════════════════════

© 2026 - Sist_Prev | Sistema de Planejamento Previdenciário
Desenvolvido para advogados previdenciários

Versão 2.0 - Janeiro 2026
- Extração 100% completa (178/178 remunerações)
- Arquitetura modular e testável
- Pronto para próxima fase: Motor de Cálculo

═══════════════════════════════════════════════════════════════════════════════════════
