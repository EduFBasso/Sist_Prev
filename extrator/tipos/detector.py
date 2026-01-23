"""
Detector de Tipos de Vínculos CNIS

Identifica o tipo de vínculo/contribuição analisando marcadores no texto do PDF.

TIPOS SUPORTADOS:
1. CLT (Empregado/Agente Público)
   - Marcador: "Matrícula do Tipo Filiado" + "Código Emp."
   - Seção: "Remunerações"
   
2. FACULTATIVO (Contribuinte Individual)
   - Marcador: "Origem do Vínculo" + NIT + "RECOLHIMENTO"
   - NÃO possui: "Código Emp."
   - Seção: "Contribuições"

3. AUTÔNOMO/EMPRESÁRIO (futuro)
4. SEGURADO ESPECIAL (futuro)
5. MEI (futuro)
"""

import re
from enum import Enum, auto


class TipoVinculo(Enum):
    """Tipos de vínculos previdenciários suportados."""
    CLT = auto()
    FACULTATIVO = auto()
    AUTONOMO = auto()
    EMPRESARIO = auto()
    SEGURADO_ESPECIAL = auto()
    MEI = auto()
    DESCONHECIDO = auto()


def detectar_tipo_vinculo(texto_bloco: str) -> TipoVinculo:
    """Detecta o tipo de vínculo analisando marcadores no texto.
    
    Args:
        texto_bloco: Texto do bloco de vínculo (entre "Matrícula" e próximo vínculo)
        
    Returns:
        TipoVinculo identificado
        
    Examples:
        >>> texto_clt = "Matrícula do Tipo Filiado\\nSeq. NIT Código Emp.\\n1 125.xxx CNPJ..."
        >>> detectar_tipo_vinculo(texto_clt)
        TipoVinculo.CLT
        
        >>> texto_fac = "Origem do Vínculo\\n11 125.xxx RECOLHIMENTO..."
        >>> detectar_tipo_vinculo(texto_fac)
        TipoVinculo.FACULTATIVO
    """
    
    # CLT: Possui "Código Emp." (CNPJ da empresa)
    if "Código Emp." in texto_bloco or "Codigo Emp." in texto_bloco:
        return TipoVinculo.CLT
    
    # FACULTATIVO: "Origem do Vínculo" + NIT + "RECOLHIMENTO" sem CNPJ
    if "Origem do Vínculo" in texto_bloco or "Origem do Vinculo" in texto_bloco:
        # Verificar se tem NIT pattern (xxx.xxxxx.xx-x)
        if re.search(r'\d{3}\.\d{5}\.\d{2}-\d', texto_bloco):
            # Verificar se menciona RECOLHIMENTO/FACULTATIVO
            if any(x in texto_bloco.upper() for x in ["RECOLHIMENTO", "FACULTATIVO", "CONTRIBUINTE INDIVIDUAL"]):
                return TipoVinculo.FACULTATIVO
    
    # MEI: Microempreendedor Individual (futuro)
    if "MEI" in texto_bloco or "MICROEMPREENDEDOR" in texto_bloco.upper():
        return TipoVinculo.MEI
    
    # AUTÔNOMO: Contribuinte Individual com CNPJ próprio (futuro)
    if "CONTRIBUINTE INDIVIDUAL" in texto_bloco.upper() and "CNPJ" in texto_bloco:
        return TipoVinculo.AUTONOMO
    
    # SEGURADO ESPECIAL: Regime especial (futuro)
    if "SEGURADO ESPECIAL" in texto_bloco.upper():
        return TipoVinculo.SEGURADO_ESPECIAL
    
    return TipoVinculo.DESCONHECIDO


def detectar_tipo_vinculo_por_seq(seq: str, nit: str, codigo_emp: str, 
                                   texto_contexto: str = "") -> TipoVinculo:
    """Detecta tipo de vínculo pelos dados estruturados (Seq, NIT, CNPJ).
    
    Útil quando já temos os campos extraídos (não precisa analisar texto bruto).
    
    Args:
        seq: Número sequência do vínculo
        nit: NIT do segurado
        codigo_emp: CNPJ da empresa ou "FACULTATIVO"
        texto_contexto: Texto adicional para análise (opcional)
        
    Returns:
        TipoVinculo identificado
    """
    
    # FACULTATIVO: codigo_emp = "FACULTATIVO" ou sem CNPJ válido
    if codigo_emp == "FACULTATIVO" or not codigo_emp or not re.match(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}', codigo_emp):
        return TipoVinculo.FACULTATIVO
    
    # CLT: Possui CNPJ válido
    if re.match(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}', codigo_emp):
        return TipoVinculo.CLT
    
    # Análise adicional pelo contexto se disponível
    if texto_contexto:
        return detectar_tipo_vinculo(texto_contexto)
    
    return TipoVinculo.DESCONHECIDO


def obter_nome_tipo(tipo: TipoVinculo) -> str:
    """Retorna nome legível do tipo de vínculo.
    
    Args:
        tipo: Enum TipoVinculo
        
    Returns:
        String com nome legível
    """
    nomes = {
        TipoVinculo.CLT: "CLT (Empregado/Agente Público)",
        TipoVinculo.FACULTATIVO: "Facultativo (Contribuinte Individual)",
        TipoVinculo.AUTONOMO: "Autônomo (Contribuinte Individual com CNPJ)",
        TipoVinculo.EMPRESARIO: "Empresário",
        TipoVinculo.SEGURADO_ESPECIAL: "Segurado Especial",
        TipoVinculo.MEI: "MEI (Microempreendedor Individual)",
        TipoVinculo.DESCONHECIDO: "Desconhecido",
    }
    return nomes.get(tipo, "Tipo não identificado")
