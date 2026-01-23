"""
Módulo Extrator de CNIS - Estrutura Modular v2.0

Este módulo contém a lógica de extração de dados do CNIS (PDF → CSV).

ARQUITETURA MODULAR:
├── tipos/                    # Processadores por tipo de vínculo
│   ├── clt.py               # CLT (Empregado) - 3 campos
│   ├── facultativo.py       # Facultativo - 5 campos  
│   └── detector.py          # Detecta tipo de vínculo
├── parsers/                 # Parsers especializados
│   ├── vinculos.py          # Parse de vínculos
│   └── cabecalho.py         # Parse de cabeçalho
├── orquestrador.py          # Coordena extração completa
└── utils.py                 # Funções auxiliares

FILOSOFIA:
- Cada tipo de vínculo/contribuição em módulo separado
- Prioridade: Modularização (separação clara)
- Próxima fase: Reutilização de código comum
"""

__version__ = "2.0.0"
__author__ = "Sistema Prev"

# Imports públicos - tipos
from .tipos import (
    processar_remuneracoes_clt,
    processar_valores_soltos_clt,
    processar_contribuicoes_facultativo,
    detectar_tipo_vinculo,
    TipoVinculo,
)

# Imports públicos - parsers
from .parsers import (
    extrair_vinculos_texto,
    parse_vinculo_texto,
    salvar_vinculos_estruturados,
    extrair_dados_cabecalho,
    salvar_cabecalho_csv,
)

# Imports públicos - utils
from .utils import (
    detectar_encoding,
    limpar_remuneracao,
    validar_valor,
    extrair_zona_util,
)

# Import público - orquestrador
from .orquestrador import processar_cnis_completo

__all__ = [
    # Tipos
    'processar_remuneracoes_clt',
    'processar_valores_soltos_clt',
    'processar_contribuicoes_facultativo',
    'detectar_tipo_vinculo',
    'TipoVinculo',
    # Parsers
    'extrair_vinculos_texto',
    'parse_vinculo_texto',
    'salvar_vinculos_estruturados',
    'extrair_dados_cabecalho',
    'salvar_cabecalho_csv',
    # Utils
    'detectar_encoding',
    'limpar_remuneracao',
    'validar_valor',
    'extrair_zona_util',
    # Orquestrador
    'processar_cnis_completo',
]
