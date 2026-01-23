"""
Parser de Vínculos - Extração e Parse de Vínculos CLT

Extrai blocos de vínculos diretamente do texto e tabelas do PDF.
"""

import csv
import re
from pathlib import Path


def extrair_vinculos_texto(caminho_pdf: str):
    """Extrai blocos de vínculos diretamente do texto das páginas do PDF.

    Complementa a extração por tabelas, capturando vínculos que não
    foram convertidos em tabela.
    
    Args:
        caminho_pdf: Caminho do arquivo PDF
        
    Returns:
        Lista de tuplas (pagina_idx, texto_bloco)
    """
    import pdfplumber
    
    pdf_path = Path(caminho_pdf)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF não encontrado: {pdf_path}")

    blocos: list[tuple[int, str]] = []

    with pdfplumber.open(pdf_path) as pdf:
        for pagina_idx, pagina in enumerate(pdf.pages, start=1):
            texto = pagina.extract_text() or ""
            marcador = "Matrícula do Tipo Filiado no"
            pos = 0

            while True:
                inicio = texto.find(marcador, pos)
                if inicio == -1:
                    break

                # Bloco termina antes de "Remunerações"
                fim = texto.find("Remunerações", inicio)
                if fim == -1:
                    fim = len(texto)

                bloco = texto[inicio:fim]
                blocos.append((pagina_idx, bloco))

                pos = fim

    return blocos


def parse_vinculo_texto(texto: str) -> dict:
    """Extrai campos estruturados de um bloco de vínculo CLT do CNIS.

    Formato esperado:
        "Matrícula do Tipo Filiado no\\n"
        "Seq. NIT Código Emp. Origem do Vínculo...\\n"
        "1 125.37781.66-1 56.528.946/0001-80 EMPRESA XYZ LTDA Empregado ou Agente...\\n"

    Returns:
        Dicionário com: seq, nit, codigo_emp, empresa, tipo_filiado, 
                       data_inicio, data_fim, ult_remun
        Ou dicionário vazio se falhar
    """
    try:
        linhas = str(texto).splitlines()
        # Ignora as 2 primeiras linhas de cabeçalho
        corpo = " ".join(linhas[2:]).strip()
        # Normaliza espaços
        corpo = " ".join(corpo.split())

        marcador_tipo = " Empregado ou Agente "
        if marcador_tipo not in corpo:
            return {}

        antes, depois = corpo.split(marcador_tipo, 1)
        tipo_filiado = "Empregado ou Agente"

        partes_antes = antes.split()
        if len(partes_antes) < 3:
            return {}

        seq = partes_antes[0]
        nit = partes_antes[1]
        codigo_emp = partes_antes[2]

        # Empresa: tudo entre CNPJ e tipo_filiado
        empresa_inicio = antes.find(codigo_emp) + len(codigo_emp)
        empresa = antes[empresa_inicio:].strip()

        # Datas e última remuneração
        partes_depois = depois.split()
        data_inicio = ""
        data_fim = ""
        ult_remun = ""

        if len(partes_depois) >= 1:
            for parte in partes_depois:
                if re.match(r'\\d{2}/\\d{2}/\\d{4}', parte):
                    if not data_inicio:
                        data_inicio = parte
                    elif not data_fim:
                        data_fim = parte
                elif re.match(r'\\d{2}/\\d{4}', parte):
                    ult_remun = parte

        return {
            "seq": seq,
            "nit": nit,
            "codigo_emp": codigo_emp,
            "empresa": empresa,
            "tipo_filiado": tipo_filiado,
            "data_inicio": data_inicio,
            "data_fim": data_fim,
            "ult_remun": ult_remun,
        }

    except Exception:
        return {}


def salvar_vinculos_estruturados(linhas_saida, caminho_csv: str, 
                                 dados_cab: dict | None = None, 
                                 caminho_pdf: str | None = None) -> None:
    """Gera um CSV estruturado com um vínculo por linha.

    Consolida vínculos de tabelas + texto, remove duplicatas, ordena por Seq.
    
    Args:
        linhas_saida: Linhas das tabelas extraídas
        caminho_csv: Caminho do arquivo CSV de saída
        dados_cab: Dados do cabeçalho (cliente)
        caminho_pdf: Caminho do PDF (para extração complementar de texto)
    """
    csv_path = Path(caminho_csv)

    # 1) Coletar vínculos vindos das tabelas
    registros: list[dict] = []

    for linha in linhas_saida:
        pagina_idx, tabela_idx, *cols = linha
        for col in cols:
            if "Matrícula do Tipo Filiado no" in str(col):
                dados = parse_vinculo_texto(str(col))
                if dados:
                    registros.append({
                        "pagina": pagina_idx,
                        "tabela": tabela_idx,
                        **dados,
                    })
                break

    # 2) Complementar com vínculos extraídos do texto (se fornecido)
    if caminho_pdf:
        for pagina_idx, bloco in extrair_vinculos_texto(caminho_pdf):
            dados = parse_vinculo_texto(bloco)
            if dados:
                registros.append({
                    "pagina": pagina_idx,
                    "tabela": 0,
                    **dados,
                })

    # 3) Remover duplicidades por (Seq, NIT, CodigoEmp)
    vistos: set[tuple[str, str, str]] = set()
    unicos: list[dict] = []

    for r in registros:
        chave = (str(r.get("seq", "")), str(r.get("nit", "")), str(r.get("codigo_emp", "")))
        if chave in vistos:
            continue
        vistos.add(chave)
        unicos.append(r)
    
    # Ordenar por Seq (numérico)
    unicos_ordenados = sorted(unicos, key=lambda x: int(x.get("seq", "0")) if str(x.get("seq", "0")).isdigit() else 0)

    # 4) Gravar CSV final estruturado
    with csv_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=";")

        header = [
            "Pagina",
            "Tabela",
            "Seq",
            "NIT",
            "CodigoEmp",
            "Empresa",
            "TipoFiliado",
            "DataInicio",
            "DataFim",
            "UltRemunCompetencia",
        ]

        # Incluir dados do cliente se disponíveis
        if dados_cab:
            header.extend([
                "NIT_Cliente",
                "CPF_Cliente",
                "NomeCliente",
                "DataNascimentoCliente",
                "NomeMaeCliente",
            ])

        writer.writerow(header)

        for r in unicos_ordenados:
            linha_saida = [
                r.get("pagina", ""),
                r.get("tabela", ""),
                r.get("seq", ""),
                r.get("nit", ""),
                r.get("codigo_emp", ""),
                r.get("empresa", ""),
                r.get("tipo_filiado", ""),
                r.get("data_inicio", ""),
                r.get("data_fim", ""),
                r.get("ult_remun", ""),
            ]

            if dados_cab:
                linha_saida.extend([
                    dados_cab.get("NIT", ""),
                    dados_cab.get("CPF", ""),
                    dados_cab.get("Nome", ""),
                    dados_cab.get("DataNascimento", ""),
                    dados_cab.get("NomeMae", ""),
                ])

            writer.writerow(linha_saida)
