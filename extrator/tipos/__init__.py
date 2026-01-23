"""
Módulos de Tipos de Vínculos CNIS

Cada tipo de vínculo/contribuição tem seu próprio módulo especializado.
"""

from .clt import processar_remuneracoes_clt, processar_valores_soltos_clt
from .facultativo import processar_contribuicoes_facultativo
from .detector import detectar_tipo_vinculo, detectar_tipo_vinculo_por_seq, obter_nome_tipo, TipoVinculo

__all__ = [
    'processar_remuneracoes_clt',
    'processar_valores_soltos_clt',
    'processar_contribuicoes_facultativo',
    'detectar_tipo_vinculo',
    'detectar_tipo_vinculo_por_seq',
    'obter_nome_tipo',
    'TipoVinculo',
]
