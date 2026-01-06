import sys
import csv
from pathlib import Path

import pdfplumber


def extrair_tabelas_para_csv(caminho_pdf: str, caminho_csv: str) -> None:
    pdf_path = Path(caminho_pdf)
    csv_path = Path(caminho_csv)

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

    # Descobrir o maior número de colunas encontradas
    max_cols = 0
    for linha in linhas_saida:
        max_cols = max(max_cols, len(linha) - 2)  # desconsidera página/tabela

    # Cabeçalho genérico: Page,Table,Col1,Col2,...
    header = ["Pagina", "Tabela"] + [f"Col{i}" for i in range(1, max_cols + 1)]

    with csv_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(header)
        for linha in linhas_saida:
            pagina_idx, tabela_idx, *cols = linha
            # Completa com strings vazias até max_cols
            cols = (cols + [""] * max_cols)[:max_cols]
            writer.writerow([pagina_idx, tabela_idx, *cols])


def main(argv=None) -> None:
    if argv is None:
        argv = sys.argv[1:]

    if len(argv) < 2:
        print("Uso: python converter_extrato_inss.py <entrada.pdf> <saida.csv>")
        sys.exit(1)

    pdf_in = argv[0]
    csv_out = argv[1]

    extrair_tabelas_para_csv(pdf_in, csv_out)
    print(f"Arquivo CSV gerado em: {csv_out}")


if __name__ == "__main__":
    main()
