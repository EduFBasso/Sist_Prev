"""
Configurações dos tipos de vínculos - Padrões de captura

Define as características de cada tipo de vínculo para facilitar 
a manutenção e extensão do código.
"""

from dataclasses import dataclass
from typing import Optional
from .detector import TipoVinculo


@dataclass
class ConfigVinculo:
    """Configuração de padrões de captura para um tipo de vínculo."""
    
    # Identificação
    tipo: TipoVinculo
    nome: str
    
    # Marcadores de início do bloco
    marcador_bloco_clt: Optional[str] = None  # Ex: "Matrícula do Tipo Filiado"
    marcador_bloco_facultativo: Optional[str] = None  # Ex: "Seq. NIT Origem do Vínculo"
    
    # Padrões de detecção
    padrao_deteccao_cnpj: bool = False  # True se usa CNPJ (CLT)
    padrao_deteccao_texto: Optional[str] = None  # Ex: "RECOLHIMENTO Facultativo"
    
    # Nome da seção de dados
    nome_secao: str = "Remunerações"  # "Remunerações" ou "Contribuições"
    
    # Padrões de cabeçalho da seção
    cabecalho_completo: str = ""  # Ex: "Remunerações\s+Competência\s+Remuneração"
    cabecalho_minimo: str = ""  # Ex: "Competência\s+Remuneração\s+Indicadores"
    
    # Estrutura dos dados
    num_campos: int = 3  # 3 campos (CLT) ou 5 campos (Facultativo)
    tem_data_pagamento: bool = False  # True para Facultativo
    
    # Identificador de código empresa
    codigo_emp_padrao: str = ""  # "FACULTATIVO", "MEI", etc. (se não usa CNPJ)
    usa_cnpj: bool = True  # False para Facultativo, MEI, etc.
    
    # Processamento de blocos cortados
    suporta_corte_rodape: bool = True  # Se pode ser cortado pelo rodapé


# ============================================================================
# CONFIGURAÇÕES POR TIPO
# ============================================================================

CONFIG_CLT = ConfigVinculo(
    tipo=TipoVinculo.CLT,
    nome="CLT",
    marcador_bloco_clt="Matrícula do Tipo Filiado",
    padrao_deteccao_cnpj=True,
    nome_secao="Remunerações",
    cabecalho_completo=r'Remunerações\s+Competência\s+Remuneração',
    cabecalho_minimo=r'Competência\s+Remuneração\s+Indicadores',
    num_campos=3,
    tem_data_pagamento=False,
    codigo_emp_padrao="",  # Usa CNPJ
    usa_cnpj=True,
    suporta_corte_rodape=True
)

CONFIG_FACULTATIVO = ConfigVinculo(
    tipo=TipoVinculo.FACULTATIVO,
    nome="FACULTATIVO",
    marcador_bloco_facultativo="Seq. NIT Origem do Vínculo",
    padrao_deteccao_texto="Facultativo",
    nome_secao="Contribuições",
    cabecalho_completo=r'Contribuições\s+Competência\s+Data\s+Pgto',
    cabecalho_minimo=r'Competência\s+Data\s+Pgto',
    num_campos=5,
    tem_data_pagamento=True,
    codigo_emp_padrao="FACULTATIVO",
    usa_cnpj=False,
    suporta_corte_rodape=True
)

CONFIG_MEI = ConfigVinculo(
    tipo=TipoVinculo.MEI,
    nome="MEI",
    marcador_bloco_facultativo="Seq. NIT Origem do Vínculo",
    padrao_deteccao_texto="MEI",
    nome_secao="Contribuições",
    cabecalho_completo=r'Contribuições\s+Competência\s+Data\s+Pgto',
    cabecalho_minimo=r'Competência\s+Data\s+Pgto',
    num_campos=5,
    tem_data_pagamento=True,
    codigo_emp_padrao="MEI",
    usa_cnpj=False,
    suporta_corte_rodape=True
)

CONFIG_AUTONOMO = ConfigVinculo(
    tipo=TipoVinculo.AUTONOMO,
    nome="AUTONOMO",
    marcador_bloco_facultativo="Seq. NIT Origem do Vínculo",
    padrao_deteccao_texto="Contribuinte Individual",
    nome_secao="Contribuições",
    cabecalho_completo=r'Contribuições\s+Competência\s+Data\s+Pgto',
    cabecalho_minimo=r'Competência\s+Data\s+Pgto',
    num_campos=5,
    tem_data_pagamento=True,
    codigo_emp_padrao="AUTONOMO",
    usa_cnpj=False,
    suporta_corte_rodape=True
)


# ============================================================================
# MAPEAMENTO TIPO → CONFIGURAÇÃO
# ============================================================================

CONFIGS_POR_TIPO = {
    TipoVinculo.CLT: CONFIG_CLT,
    TipoVinculo.FACULTATIVO: CONFIG_FACULTATIVO,
    TipoVinculo.MEI: CONFIG_MEI,
    TipoVinculo.AUTONOMO: CONFIG_AUTONOMO,
}


def obter_config(tipo: TipoVinculo) -> Optional[ConfigVinculo]:
    """Obtém a configuração para um tipo de vínculo."""
    return CONFIGS_POR_TIPO.get(tipo)


def obter_marcadores_blocos() -> tuple[str, str]:
    """
    Retorna os marcadores de início de blocos.
    
    Returns:
        (marcador_clt, marcador_facultativo): Tupla com os dois marcadores principais
    """
    return (
        CONFIG_CLT.marcador_bloco_clt,
        CONFIG_FACULTATIVO.marcador_bloco_facultativo
    )


def obter_tipos_com_contribuicoes() -> list[TipoVinculo]:
    """
    Retorna lista de tipos que usam seção "Contribuições" (não "Remunerações").
    """
    return [
        tipo for tipo, config in CONFIGS_POR_TIPO.items()
        if config.nome_secao == "Contribuições"
    ]


def obter_tipos_sem_cnpj() -> list[TipoVinculo]:
    """
    Retorna lista de tipos que não usam CNPJ (facultativo, MEI, autônomo).
    """
    return [
        tipo for tipo, config in CONFIGS_POR_TIPO.items()
        if not config.usa_cnpj
    ]
