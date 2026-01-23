"""
Módulo Extrator de CNIS

Este módulo contém a lógica de extração de dados do CNIS (PDF → CSV).

Estrutura:
- extrator_clt.py: Extração de Remunerações (CLT, 3 campos)
- extrator_facultativo.py: Extração de Contribuições (Facultativo, 5 campos)
- parser_vinculos.py: Parse de vínculos (tabela + texto)
- utils.py: Funções auxiliares (encoding, limpeza, etc)
"""

__version__ = "2.0.0"
__author__ = "Sistema Prev"

# Imports para facilitar uso externo
from .extrator_clt import processar_remuneracoes
from .extrator_facultativo import processar_contribuicoes_facultativo
from .parser_vinculos import (
    extrair_vinculos_texto,
    parse_vinculo_texto,
    salvar_vinculos_estruturados
)
from .utils import (
    detectar_encoding, 
    limpar_remuneracao, 
    validar_valor, 
    extrair_zona_util, 
    extrair_dados_cabecalho
)

__all__ = [
    'processar_remuneracoes',
    'processar_contribuicoes_facultativo',
    'extrair_vinculos_texto',
    'parse_vinculo_texto',
    'salvar_vinculos_estruturados',
    'detectar_encoding',
    'limpar_remuneracao',
    'validar_valor',
    'extrair_zona_util',
    'extrair_dados_cabecalho',
]
