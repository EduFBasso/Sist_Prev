"""
Parsers Especializados

Parsers para diferentes seções do CNIS (vínculos, cabeçalho, etc).
"""

from .vinculos import extrair_vinculos_texto, parse_vinculo_texto, salvar_vinculos_estruturados
from .cabecalho import extrair_dados_cabecalho, salvar_cabecalho_csv

__all__ = [
    'extrair_vinculos_texto',
    'parse_vinculo_texto',
    'salvar_vinculos_estruturados',
    'extrair_dados_cabecalho',
    'salvar_cabecalho_csv',
]
