"""
Módulos de Tipos de Vínculos CNIS

Cada tipo de vínculo/contribuição tem seu próprio módulo especializado.
"""

# Detector de tipo de documento CNIS
from .tipo_cnis import TipoCNIS, detectar_tipo_cnis, obter_nome_tipo_cnis, validar_tipo_cnis

# Detector de tipo de vínculo
from .detector import detectar_tipo_vinculo, detectar_tipo_vinculo_por_seq, obter_nome_tipo, TipoVinculo

# Processadores por tipo de vínculo
from .clt import processar_remuneracoes_clt, processar_valores_soltos_clt
from .facultativo import processar_contribuicoes_facultativo

# Coordenador de remunerações (delegação automática)
from .coordenador_remuneracoes import (
    extrair_remuneracoes_coordenado,
    extrair_remuneracoes_texto,  # compatibilidade com legado
    validar_baseline_remuneracoes
)

__all__ = [
    'processar_remuneracoes_clt',
    'processar_valores_soltos_clt',
    'processar_contribuicoes_facultativo',
    'detectar_tipo_vinculo',
    'detectar_tipo_vinculo_por_seq',
    'obter_nome_tipo',
    'TipoVinculo',
]
