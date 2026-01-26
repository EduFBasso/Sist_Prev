#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Conversor de Extrato CNIS (PDF) para CSV

Este script extrai dados do CNIS (Cadastro Nacional de Informações Sociais) do INSS
em formato PDF e converte para arquivos CSV estruturados.

SUPORTE A FORMATOS:
1. CLT (Vínculos com Empregador - Seq 1-10)
   - Detectado por: "Matrícula do Tipo Filiado" + "Código Emp." presente
   - Seção: "Remunerações"
   - Formato: Competência | Remuneração | Indicadores (3 campos)
   - Até 3 competências por linha

2. FACULTATIVO (Contribuinte Facultativo - Seq 11+)
   - Detectado por: "Origem do Vínculo" + NIT pattern + "RECOLHIMENTO" + sem "Código Emp."
   - Seção: "Contribuições"
   - Formato: Competência | Data Pagto | Contribuição | Salário | Indicadores (5 campos)
   - Captura: Competência, Contribuição, Indicadores (ignora Data Pagto e Salário)
   - Até 2 competências por linha

ARQUIVOS GERADOS:
- [base].csv                          : Dados brutos das tabelas
- [base]_dados_cliente.csv            : Identificação do filiado
- [base]_vinculos_brutos.csv          : Blocos de vínculos (texto)
- [base]_vinculos_estruturado.csv     : Vínculos parseados (estruturado)
- [base]_remuneracoes.csv             : Remunerações/Contribuições

USO:
    python converter_extrato_inss.py <caminho_pdf> <caminho_csv_saida>

EXEMPLO:
    python converter_extrato_inss.py cnis/CNIS_JOAO.pdf saida/cnis.csv

AUTOR: Sistema Prev
DATA: 2026-01-23
"""

import sys
import csv
import re
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


def extrair_vinculos_texto(caminho_pdf: str):
    """Extrai blocos de vínculos diretamente do texto das páginas do PDF.

    Isso complementa a extração por tabelas, permitindo capturar vínculos
    que não foram convertidos em tabela (como a Seq. 2 do seu exemplo).
    Retorna uma lista de tuplas (pagina_idx, texto_bloco).
    """

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

                # Em geral o bloco de vínculo termina logo antes de "Remunerações"
                fim = texto.find("Remunerações", inicio)
                if fim == -1:
                    fim = len(texto)

                bloco = texto[inicio:fim]
                blocos.append((pagina_idx, bloco))

                pos = fim

    return blocos


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
        # Normaliza espaços em branco (sem alterar a ordem dos tokens)
        corpo = " ".join(corpo.split())

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
        if not partes_depois:
            return {}

        # Identificação de datas completas (dd/mm/aaaa) e competência (mm/aaaa)
        def is_data_completa(token: str) -> bool:
            return re.fullmatch(r"\d{2}/\d{2}/\d{4}", token) is not None

        def is_competencia(token: str) -> bool:
            return re.fullmatch(r"\d{2}/\d{4}", token) is not None

        # Mantemos apenas os tokens que parecem datas/competências, na ordem em que aparecem
        datas = [t for t in partes_depois if is_data_completa(t) or is_competencia(t)]
        if not datas:
            return {}

        data_inicio = ""
        data_fim = ""
        ult_remun = ""

        if len(datas) == 1:
            # Só uma data encontrada: consideramos como Data Início
            data_inicio = datas[0]
        elif len(datas) == 2:
            # Duas datas:
            #  - se forem (dd/mm/aaaa, dd/mm/aaaa) -> início e fim (sem última remuneração)
            #  - se forem (dd/mm/aaaa, mm/aaaa)    -> início e última remuneração (fim em branco)
            if is_data_completa(datas[0]) and is_data_completa(datas[1]):
                data_inicio = datas[0]
                data_fim = datas[1]
            else:
                data_inicio = datas[0]
                ult_remun = datas[1]
        else:
            # Três ou mais datas: padrão típico do CNIS
            # 1) Data Início (dd/mm/aaaa)
            # 2) Data Fim   (dd/mm/aaaa)
            # 3) Últ. Remun (mm/aaaa)
            data_inicio = datas[0]
            data_fim = datas[1]
            ult_remun = datas[2]

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


def salvar_vinculos_estruturados(linhas_saida, caminho_csv: str, dados_cab: dict | None = None, caminho_pdf: str | None = None) -> None:
    """Gera um CSV estruturado com um vínculo por linha.

    Este arquivo já se aproxima bastante do formato que o VBA
    poderá importar para a aba Vinculos.
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

    # 2) Complementar com vínculos extraídos diretamente do texto do PDF (se fornecido)
    if caminho_pdf:
        for pagina_idx, bloco in extrair_vinculos_texto(caminho_pdf):
            dados = parse_vinculo_texto(bloco)
            if dados:
                registros.append({
                    "pagina": pagina_idx,
                    "tabela": 0,
                    **dados,
                })

    # 3) Remover duplicidades (mesmo vínculo vindo de tabela e de texto)
    #    Chave de deduplicação: (Seq, NIT, CodigoEmp)
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

        # Inclui também os dados do cliente (cabeçalho) se disponíveis
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


def extrair_dados_cabecalho(caminho_pdf: str) -> dict:
    """Extrai dados do cabeçalho do extrato (Identificação do Filiado).

    Tenta obter: NIT, CPF, Nome, DataNascimento, NomeMae.
    """

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
    """Salva os dados do cabeçalho em um pequeno CSV (uma linha)."""

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


def processar_contribuicoes_facultativo(secao_contrib: str, seq: str, pagina_idx: int, 
                                       registros: list) -> None:
    """Processa seção de CONTRIBUIÇÕES (vínculos Facultativos).
    
    FORMATO FACULTATIVO (5 campos):
        Contribuições
        Competência   Data Pagto.   Contribuição   Salário Contrib.   Indicadores
        09/2019       15/09/2019    200,00         1045,00            PREC-FACULTCONC
        10/2019       15/10/2019    199,60         1045,00            PREC-FACULTCONC
    
    REGEX PATTERN:
        (\d{2}/\d{4})                      # Competência (captura)
        \s+\d{2}/\d{2}/\d{4}               # Data Pagto (ignora)
        \s+([\d.,]+)                       # Contribuição (captura como remuneracao)
        \s+([\d.,]+)                       # Salário (ignora)
        \s*([^\d/]*?)                      # Indicadores (captura)
        (?=\d{2}/\d{4}|$)                  # Lookahead próxima competência ou fim
    
    DIFERENÇAS vs CLT:
        - 5 campos vs 3 campos
        - Até 2 competências por linha vs 3
        - codigo_emp = "FACULTATIVO" (sem CNPJ)
        - Usa "Contribuição" como valor de remuneracao
    
    Args:
        secao_contrib: Texto da seção "Contribuições" até próximo vínculo
        seq: Número de sequência do vínculo (ex: "11")
        pagina_idx: Número da página do PDF (1-indexed)
        registros: Lista onde adicionar dicionários de remunerações
    """
    linhas = secao_contrib.split('\n')
    
    # Pular cabeçalho
    i = 0
    while i < len(linhas) and not re.search(r'\d{2}/\d{4}', linhas[i]):
        i += 1
    
    # Processar linhas de dados
    while i < len(linhas):
        linha = linhas[i].strip()
        
        # Parar se encontrar novo vínculo
        if not linha or any(x in linha for x in ["Matrícula", "Seq.", "Origem do Vínculo"]):
            break
        
        # Regex: Competência + Data Pagto (dd/mm/aaaa) + Contribuição + Salário + Indicadores
        # Exemplo: 09/2019 18/09/2019 200,00 1.000,00 PREC-FACULTCONC
        padroes = re.findall(
            r'(\d{2}/\d{4})\s+\d{2}/\d{2}/\d{4}\s+([\d.,]+)\s+([\d.,]+)\s*([^\d/]*?)(?=\d{2}/\d{4}|$)',
            linha
        )
        
        if padroes:
            for competencia, contribuicao, salario_contrib, indicadores in padroes:
                contribuicao_limpa = contribuicao.replace('.', '').replace(',', '.')
                
                try:
                    valor = float(contribuicao_limpa)
                    if 0 < valor < 10000000:
                        registros.append({
                            "pagina": pagina_idx,
                            "seq": seq,
                            "codigo_emp": "FACULTATIVO",
                            "competencia": competencia,
                            "remuneracao": contribuicao_limpa,  # Usar contribuição como remuneração
                            "indicadores": indicadores.strip() or "PREC-FACULTCONC",
                        })
                except ValueError:
                    pass
        
        i += 1


def salvar_remuneracoes_csv(caminho_pdf: str, linhas_saida, caminho_csv: str) -> None:
    """Gera um CSV com todas as remunerações extraídas por vínculo.
    
    Utiliza o coordenador modular que delega automaticamente para:
    - CLT: tipos/clt.py (3 campos)
    - FACULTATIVO: tipos/facultativo.py (5 campos)
    """
    from extrator.tipos.coordenador_remuneracoes import extrair_remuneracoes_coordenado
    
    csv_path = Path(caminho_csv)
    
    # Extrair remunerações usando o coordenador modular
    remuneracoes = extrair_remuneracoes_coordenado(caminho_pdf)
    
    # Ordenar por seq e competencia
    remuneracoes.sort(key=lambda x: (x.get("seq", ""), x.get("competencia", "")))
    
    # Salvar CSV
    with csv_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["Pagina", "Seq", "CodigoEmp", "Competencia", "Remuneracao", "Indicadores"])
        
        for r in remuneracoes:
            writer.writerow([
                r.get("pagina", ""),
                r.get("seq", ""),
                r.get("cnpj", ""),  # coordenador usa 'cnpj' em vez de 'codigo_emp'
                r.get("competencia", ""),
                r.get("remuneracao", ""),
                r.get("indicadores", ""),
            ])


def main(argv=None) -> None:
    """Função principal do CLI."""
    if argv is None:
        argv = sys.argv[1:]

    # Validação e ajuda
    if len(argv) == 0 or (len(argv) == 1 and argv[0] in ['-h', '--help', 'help']):
        print(__doc__)
        sys.exit(0)
    
    if len(argv) < 2:
        print("❌ ERRO: Argumentos insuficientes\n")
        print("Uso: python converter_extrato_inss.py <entrada.pdf> <saida.csv>\n")
        print("Exemplo: python converter_extrato_inss.py cnis/CNIS_JOAO.pdf saida/cnis.csv")
        sys.exit(1)

    pdf_in = argv[0]
    csv_out = argv[1]
    
    # Validar existência do PDF
    if not Path(pdf_in).exists():
        print(f"❌ ERRO: Arquivo não encontrado: {pdf_in}")
        sys.exit(1)

    # Preparar caminhos de saída
    base = Path(csv_out).stem
    pasta = Path(csv_out).parent
    pasta.mkdir(parents=True, exist_ok=True)
    
    # Extrair tabelas do PDF
    print(f"📄 Processando: {pdf_in}")
    linhas_saida, max_cols = extrair_tabelas(pdf_in)
    
    # Gerar arquivos CSV
    salvar_raw_csv(linhas_saida, max_cols, csv_out)
    
    cabecalho_path = str(pasta / f"{base}_dados_cliente.csv")
    dados_cab = extrair_dados_cabecalho(pdf_in)
    salvar_cabecalho_csv(dados_cab, cabecalho_path)
    
    vinculos_brutos_path = str(pasta / f"{base}_vinculos_brutos.csv")
    salvar_vinculos_brutos(linhas_saida, vinculos_brutos_path)
    
    vinculos_struct_path = str(pasta / f"{base}_vinculos_estruturado.csv")
    salvar_vinculos_estruturados(linhas_saida, vinculos_struct_path, dados_cab, pdf_in)
    
    remuneracoes_path = str(pasta / f"{base}_remuneracoes.csv")
    salvar_remuneracoes_csv(pdf_in, linhas_saida, remuneracoes_path)
    
    # Mensagens de sucesso
    print(f"\n✅ Extração concluída! Arquivos gerados:")
    print(f"   • {csv_out}")
    print(f"   • {cabecalho_path}")
    print(f"   • {vinculos_brutos_path}")
    print(f"   • {vinculos_struct_path}")
    print(f"   • {remuneracoes_path}")


if __name__ == "__main__":
    main()

