"""
Orquestrador de Extração CNIS

Coordena a extração completa do PDF CNIS:
1. Extrai tabelas brutas
2. Identifica e extrai dados do cabeçalho
3. Detecta tipos de vínculos
4. Processa remunerações por tipo (CLT, Facultativo, etc)
5. Gera arquivos CSV estruturados

Este módulo atua como "maestro", delegando processamento específico
para os módulos especializados em cada tipo de vínculo.
"""

import csv
from pathlib import Path
from typing import List, Dict, Tuple

import pdfplumber

from .tipos import detectar_tipo_vinculo, TipoVinculo
from .tipos import processar_remuneracoes_clt, processar_contribuicoes_facultativo
from .parsers import extrair_dados_cabecalho, salvar_cabecalho_csv
from .parsers import extrair_vinculos_texto, parse_vinculo_texto, salvar_vinculos_estruturados


def extrair_tabelas_brutas(caminho_pdf: str) -> Tuple[List, int]:
    """Extrai todas as tabelas do PDF em uma lista de linhas.
    
    Args:
        caminho_pdf: Caminho do arquivo PDF
        
    Returns:
        Tupla (linhas_saida, max_cols) onde:
        - linhas_saida: Lista de linhas [pagina, tabela, col1, col2, ...]
        - max_cols: Número máximo de colunas encontrado
    """
    pdf_path = Path(caminho_pdf)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF não encontrado: {pdf_path}")
    
    linhas_saida = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for pagina_idx, pagina in enumerate(pdf.pages, start=1):
            tabelas = pagina.extract_tables()
            for tabela_idx, tabela in enumerate(tabelas, start=1):
                for linha in tabela:
                    linhas_saida.append([
                        pagina_idx,
                        tabela_idx,
                        *(cel if cel is not None else "" for cel in linha),
                    ])
    
    # Descobrir maior número de colunas
    max_cols = 0
    for linha in linhas_saida:
        max_cols = max(max_cols, len(linha) - 2)
    
    return linhas_saida, max_cols


def salvar_tabelas_raw(linhas_saida: List, max_cols: int, caminho_csv: str) -> None:
    """Salva todas as tabelas extraídas em um CSV genérico.
    
    Args:
        linhas_saida: Lista de linhas das tabelas
        max_cols: Número máximo de colunas
        caminho_csv: Caminho do arquivo CSV de saída
    """
    csv_path = Path(caminho_csv)
    
    header = ["Pagina", "Tabela"] + [f"Col{i}" for i in range(1, max_cols + 1)]
    
    with csv_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(header)
        for linha in linhas_saida:
            pagina_idx, tabela_idx, *cols = linha
            cols = (cols + [""] * max_cols)[:max_cols]
            writer.writerow([pagina_idx, tabela_idx, *cols])


def salvar_vinculos_brutos(linhas_saida: List, caminho_csv: str) -> None:
    """Gera um CSV só com os blocos de "Matrícula do Tipo Filiado".
    
    Args:
        linhas_saida: Lista de linhas das tabelas
        caminho_csv: Caminho do arquivo CSV de saída
    """
    csv_path = Path(caminho_csv)
    
    with csv_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["Pagina", "Tabela", "TextoBruto"])
        
        for linha in linhas_saida:
            pagina_idx, tabela_idx, *cols = linha
            for col in cols:
                if "Matrícula do Tipo Filiado no" in str(col):
                    texto = str(col)
                    writer.writerow([pagina_idx, tabela_idx, texto])
                    break


def processar_cnis_completo(caminho_pdf: str, pasta_saida: str, nome_base: str) -> Dict[str, str]:
    """Processa um arquivo CNIS completo e gera todos os CSVs.
    
    Esta é a função principal do orquestrador. Coordena:
    1. Extração de tabelas brutas
    2. Extração de dados do cabeçalho (identificação)
    3. Extração de vínculos (brutos e estruturados)
    4. Extração de remunerações por tipo de vínculo
    
    Args:
        caminho_pdf: Caminho do arquivo PDF do CNIS
        pasta_saida: Pasta onde salvar os CSVs gerados
        nome_base: Nome base para os arquivos (ex: "cnis_joao")
        
    Returns:
        Dicionário com caminhos dos arquivos gerados:
        {
            "raw": "path/cnis.csv",
            "cabecalho": "path/cnis_dados_cliente.csv",
            "vinculos_brutos": "path/cnis_vinculos_brutos.csv",
            "vinculos_estruturado": "path/cnis_vinculos_estruturado.csv",
            "remuneracoes": "path/cnis_remuneracoes.csv"
        }
        
    Example:
        >>> arquivos = processar_cnis_completo(
        ...     "cnis/JOAO.pdf",
        ...     "saida",
        ...     "cnis_joao"
        ... )
        >>> print(f"Remunerações: {arquivos['remuneracoes']}")
    """
    pasta = Path(pasta_saida)
    pasta.mkdir(parents=True, exist_ok=True)
    
    # 1. Extrair tabelas brutas
    linhas_saida, max_cols = extrair_tabelas_brutas(caminho_pdf)
    
    # 2. Salvar CSV raw (debug)
    csv_raw = str(pasta / f"{nome_base}.csv")
    salvar_tabelas_raw(linhas_saida, max_cols, csv_raw)
    
    # 3. Extrair e salvar dados do cabeçalho
    csv_cabecalho = str(pasta / f"{nome_base}_dados_cliente.csv")
    dados_cab = extrair_dados_cabecalho(caminho_pdf)
    salvar_cabecalho_csv(dados_cab, csv_cabecalho)
    
    # 4. Extrair e salvar vínculos brutos
    csv_vinculos_brutos = str(pasta / f"{nome_base}_vinculos_brutos.csv")
    salvar_vinculos_brutos(linhas_saida, csv_vinculos_brutos)
    
    # 5. Extrair e salvar vínculos estruturados
    csv_vinculos_struct = str(pasta / f"{nome_base}_vinculos_estruturado.csv")
    salvar_vinculos_estruturados(linhas_saida, csv_vinculos_struct, dados_cab, caminho_pdf)
    
    # 6. Extrair remunerações (usando função do converter_extrato_inss.py original)
    # TODO: Refatorar extrair_remuneracoes_texto para usar os módulos tipos/
    csv_remuneracoes = str(pasta / f"{nome_base}_remuneracoes.csv")
    # Por enquanto, retorna caminho vazio (será implementado na próxima etapa)
    
    return {
        "raw": csv_raw,
        "cabecalho": csv_cabecalho,
        "vinculos_brutos": csv_vinculos_brutos,
        "vinculos_estruturado": csv_vinculos_struct,
        "remuneracoes": csv_remuneracoes,
    }
