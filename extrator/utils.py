"""
Funções Utilitárias para Extração CNIS

Contém funções auxiliares usadas por todos os módulos.
"""

import re
from pathlib import Path


def detectar_encoding(arquivo: str) -> str:
    """Detecta encoding UTF-8 ou cp1252.
    
    Args:
        arquivo: Caminho do arquivo para detectar encoding
        
    Returns:
        'utf-8' ou 'cp1252'
    """
    try:
        with open(arquivo, 'r', encoding='utf-8') as f:
            f.read()
        return 'utf-8'
    except UnicodeDecodeError:
        return 'cp1252'


def limpar_remuneracao(valor_str: str) -> str:
    """Limpa string de remuneração para formato float.
    
    Converte "1.005,70" → "1005.70"
    
    Args:
        valor_str: String com valor no formato brasileiro (ex: "1.005,70")
        
    Returns:
        String com valor no formato float (ex: "1005.70")
    """
    return valor_str.replace('.', '').replace(',', '.')


def validar_valor(valor_float: float, min_val: float = 0, max_val: float = 10000000) -> bool:
    """Valida se valor está dentro de range aceitável.
    
    Args:
        valor_float: Valor a validar
        min_val: Valor mínimo aceitável (default: 0)
        max_val: Valor máximo aceitável (default: 10 milhões)
        
    Returns:
        True se valor válido, False caso contrário
    """
    return min_val < valor_float < max_val


def extrair_zona_util(texto_pagina: str) -> tuple[int, int]:
    """Extrai índices de início e fim da zona útil da página.
    
    Zona útil: conteúdo entre marcadores de topo e rodapé.
    
    Marcadores de topo:
    - "Relações Previdenciárias" (primeiro que encontrar)
    - "Identificação do Filiado"
    
    Marcador de rodapé:
    - "O INSS poderá rever"
    
    Args:
        texto_pagina: Texto completo da página
        
    Returns:
        Tupla (inicio_zona, fim_zona) com índices
    """
    # Buscar marcador de início
    inicio_zona = texto_pagina.find("Relações Previdenciárias")
    if inicio_zona == -1:
        inicio_zona = texto_pagina.find("Identificação do Filiado")
    if inicio_zona == -1:
        inicio_zona = 0  # Se não achar, usa início da página
    
    # Buscar marcador de fim
    fim_zona = texto_pagina.find("O INSS poderá rever")
    if fim_zona == -1:
        fim_zona = len(texto_pagina)  # Se não achar, usa fim da página
    
    return inicio_zona, fim_zona


def extrair_dados_cabecalho(caminho_pdf) -> dict:
    """Extrai dados do cabeçalho do extrato (Identificação do Filiado).
    
    Tenta obter: NIT, CPF, Nome, DataNascimento, NomeMae.
    
    Args:
        caminho_pdf: Caminho do arquivo PDF ou objeto pdfplumber.PDF
        
    Returns:
        Dicionário com dados do cliente
    """
    import pdfplumber
    
    pdf_path = Path(caminho_pdf) if isinstance(caminho_pdf, str) else None
    if pdf_path and not pdf_path.exists():
        raise FileNotFoundError(f"PDF não encontrado: {pdf_path}")
    
    if isinstance(caminho_pdf, str):
        with pdfplumber.open(caminho_pdf) as pdf:
            if not pdf.pages:
                return {}
            primeira = pdf.pages[0]
            texto = primeira.extract_text() or ""
    else:
        # Já é objeto pdfplumber.PDF
        if not caminho_pdf.pages:
            return {}
        primeira = caminho_pdf.pages[0]
        texto = primeira.extract_text() or ""
    
    dados = {
        "NIT": "",
        "CPF": "",
        "Nome": "",
        "DataNascimento": "",
        "NomeMae": "",
    }
    
    # NIT (PIS/NIT do segurado)
    m = re.search(r"NIT[:\s]+([0-9.\-]+)", texto)
    if m:
        dados["NIT"] = m.group(1)
    
    # Data de nascimento
    m = re.search(r"Data de nascimento[:\s]+(\d{2}/\d{2}/\d{4})", texto, re.IGNORECASE)
    if m:
        dados["DataNascimento"] = m.group(1)
    
    # CPF
    m = re.search(r"CPF[:\s]+([0-9.\-]+)", texto)
    if m:
        dados["CPF"] = m.group(1)
    
    # Nome
    m = re.search(r"Nome[:\s]+([A-ZÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÑÒÓÔÕÖÙÚÛÜ\s]+)", texto)
    if m:
        dados["Nome"] = m.group(1).strip()
    
    # Nome da mãe
    m = re.search(r"Nome da mãe[:\s]+([A-ZÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÑÒÓÔÕÖÙÚÛÜ\s]+)", texto, re.IGNORECASE)
    if m:
        dados["NomeMae"] = m.group(1).strip()
    
    return dados
