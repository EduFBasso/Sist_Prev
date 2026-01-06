import sys
import csv
from pathlib import Path

import pdfplumber


def extrair_tabelas(caminho_pdf: str):
    """Extrai todas as tabelas do PDF em uma lista de linhas.

    Cada linha tem o formato: [pagina, tabela, col1, col2, ...].
    Retorna (linhas_saida, max_cols).
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
                    # linha é uma lista de strings/None
                    linhas_saida.append([
                        pagina_idx,
                        tabela_idx,
                        *(cel if cel is not None else "" for cel in linha),
                    ])

    # Descobrir o maior número de colunas encontradas (sem contar página/tabela)
    max_cols = 0
    for linha in linhas_saida:
        max_cols = max(max_cols, len(linha) - 2)

    return linhas_saida, max_cols


def salvar_raw_csv(linhas_saida, max_cols: int, caminho_csv: str) -> None:
    """Salva todas as tabelas extraídas em um CSV genérico (debug/inspeção)."""

    csv_path = Path(caminho_csv)

    # Cabeçalho genérico: Pagina,Tabela,Col1,Col2,...
    header = ["Pagina", "Tabela"] + [f"Col{i}" for i in range(1, max_cols + 1)]

    with csv_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(header)
        for linha in linhas_saida:
            pagina_idx, tabela_idx, *cols = linha
            cols = (cols + [""] * max_cols)[:max_cols]
            writer.writerow([pagina_idx, tabela_idx, *cols])


def salvar_vinculos_brutos(linhas_saida, caminho_csv: str) -> None:
    """Gera um CSV só com os blocos de "Matrícula do Tipo Filiado".

    Ainda não faz o parsing completo dos campos; a ideia é ter
    um arquivo intermediário mais limpo para análise e futuros
    ajustes de parsing.
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


def parse_vinculo_texto(texto: str) -> dict:
    """Tenta extrair campos estruturados de um bloco de vínculo do CNIS.

    Espera um texto no formato aproximado de:

        "Matrícula do Tipo Filiado no\n"
        "Seq. NIT Código Emp. Origem do Vínculo Trabalhador Vínculo Data Início Data Fim Últ. Remun.\n"
        "1 125.37781.66-1 56.528.946/0001-80 EMPRESA XYZ LTDA Empregado ou Agente 19/01/1995 02/06/1995 05/1995\n"
        "Público"

    Retorna um dicionário com chaves:
      seq, nit, codigo_emp, empresa, tipo_filiado, data_inicio, data_fim, ult_remun
    Em caso de falha, retorna um dicionário vazio.
    """

    try:
        linhas = str(texto).splitlines()
        # Ignora as 2 primeiras linhas de cabeçalho
        corpo = " ".join(linhas[2:]).strip()
        corpo = " ".join(corpo.split())  # normaliza espaços

        marcador_tipo = " Empregado ou Agente "
        if marcador_tipo not in corpo:
            return {}

        antes, depois = corpo.split(marcador_tipo, 1)
        tipo_filiado = "Empregado ou Agente"

        partes_antes = antes.split()
        if len(partes_antes) < 4:
            return {}

        seq = partes_antes[0]
        nit = partes_antes[1]
        codigo_emp = partes_antes[2]
        empresa = " ".join(partes_antes[3:])

        partes_depois = depois.split()
        if len(partes_depois) < 3:
            return {}

        data_inicio = partes_depois[0]
        data_fim = partes_depois[1]
        ult_remun = partes_depois[2]

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
        # Em caso de qualquer problema, retorna vazio
        return {}


def salvar_vinculos_estruturados(linhas_saida, caminho_csv: str) -> None:
    """Gera um CSV estruturado com um vínculo por linha.

    Este arquivo já se aproxima bastante do formato que o VBA
    poderá importar para a aba Vinculos.
    """

    csv_path = Path(caminho_csv)

    with csv_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow([
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
        ])

        for linha in linhas_saida:
            pagina_idx, tabela_idx, *cols = linha
            for col in cols:
                if "Matrícula do Tipo Filiado no" in str(col):
                    dados = parse_vinculo_texto(str(col))
                    if dados:
                        writer.writerow([
                            pagina_idx,
                            tabela_idx,
                            dados["seq"],
                            dados["nit"],
                            dados["codigo_emp"],
                            dados["empresa"],
                            dados["tipo_filiado"],
                            dados["data_inicio"],
                            dados["data_fim"],
                            dados["ult_remun"],
                        ])
                    break


def main(argv=None) -> None:
    if argv is None:
        argv = sys.argv[1:]

    if len(argv) < 2:
        print("Uso: python converter_extrato_inss.py <entrada.pdf> <saida_raw.csv>")
        sys.exit(1)

    pdf_in = argv[0]
    csv_out = argv[1]

    linhas_saida, max_cols = extrair_tabelas(pdf_in)

    # CSV genérico com todas as tabelas (como antes)
    salvar_raw_csv(linhas_saida, max_cols, csv_out)

    base = Path(csv_out).stem
    pasta = Path(csv_out).parent

    # CSV adicional com os blocos de vínculos (texto bruto)
    vinculos_brutos_path = str(pasta / f"{base}_vinculos_brutos.csv")
    salvar_vinculos_brutos(linhas_saida, vinculos_brutos_path)

    # CSV estruturado com um vínculo por linha
    vinculos_struct_path = str(pasta / f"{base}_vinculos_estruturado.csv")
    salvar_vinculos_estruturados(linhas_saida, vinculos_struct_path)

    print(f"Arquivo CSV bruto gerado em: {csv_out}")
    print(f"Arquivo de vínculos (texto bruto) gerado em: {vinculos_brutos_path}")
    print(f"Arquivo de vínculos estruturados gerado em: {vinculos_struct_path}")


if __name__ == "__main__":
    main()
