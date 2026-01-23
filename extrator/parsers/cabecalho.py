"""
Parser de Cabeçalho CNIS

Extrai dados de identificação do filiado (cabeçalho do extrato).
"""

import csv
import re
from pathlib import Path


def extrair_dados_cabecalho(caminho_pdf: str) -> dict:
    """Extrai dados do cabeçalho do extrato (Identificação do Filiado).
    
    Tenta obter: NIT, CPF, Nome, DataNascimento, NomeMae.
    
    Args:
        caminho_pdf: Caminho do arquivo PDF
        
    Returns:
        Dicionário com dados do cliente
        
    Example:
        >>> dados = extrair_dados_cabecalho("cnis/JOAO.pdf")
        >>> print(dados["NIT"], dados["Nome"])
        125.37781.66-1 JOAO CARLOS EDUARDO
    """
    import pdfplumber
    
    pdf_path = Path(caminho_pdf)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF não encontrado: {pdf_path}")
    
    with pdfplumber.open(pdf_path) as pdf:
        if not pdf.pages:
            return {}
        primeira = pdf.pages[0]
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
    
    # Nome (primeira ocorrência de "Nome:" na identificação)
    m = re.search(r"Nome:\s*([^\n]+)", texto)
    if m:
        dados["Nome"] = m.group(1).strip()
    
    # Nome da mãe
    m = re.search(r"Nome da m[ãa]e:\s*([^\n]+)", texto, re.IGNORECASE)
    if m:
        dados["NomeMae"] = m.group(1).strip()
    
    if all(not v for v in dados.values()):
        return {}
    
    return dados


def salvar_cabecalho_csv(dados: dict, caminho_csv: str) -> None:
    """Salva os dados do cabeçalho em um pequeno CSV (uma linha).
    
    Args:
        dados: Dicionário com dados do cliente (NIT, CPF, Nome, etc)
        caminho_csv: Caminho do arquivo CSV de saída
    """
    if not dados:
        return
    
    csv_path = Path(caminho_csv)
    
    with csv_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["NIT", "CPF", "Nome", "DataNascimento", "NomeMae"])
        writer.writerow([
            dados.get("NIT", ""),
            dados.get("CPF", ""),
            dados.get("Nome", ""),
            dados.get("DataNascimento", ""),
            dados.get("NomeMae", ""),
        ])
