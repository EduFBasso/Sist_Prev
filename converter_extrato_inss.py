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

        for r in unicos:
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


def extrair_remuneracoes_texto(caminho_pdf: str) -> list[dict]:
    """Extrai tabelas de remuneração de cada vínculo do CNIS.
    
    Para cada vínculo, procura a seção "Remunerações" e extrai as linhas
    com formato: Competência (mm/aaaa), Remuneração (valor), Indicadores.
    
    Retorna uma lista de dicionários com:
    - seq: Sequência do vínculo
    - codigo_emp: CNPJ da empresa
    - competencia: mm/aaaa
    - remuneracao: valor
    - indicadores: texto dos indicadores (se houver)
    - pagina: número da página
    """
    
    pdf_path = Path(caminho_pdf)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF não encontrado: {pdf_path}")
    
    registros_remuneracao = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for pagina_idx, pagina in enumerate(pdf.pages, start=1):
            texto = pagina.extract_text() or ""
            
            # Procura blocos que começam com "Matrícula do Tipo Filiado"
            # e vão até o próximo vínculo ou fim da página
            marcador_vinculo = "Matrícula do Tipo Filiado"
            marcador_remuneracao = "Remunerações"
            
            pos = 0
            while True:
                inicio_vinculo = texto.find(marcador_vinculo, pos)
                if inicio_vinculo == -1:
                    break
                
                # Encontrar próximo vínculo ou fim do texto
                proximo_vinculo = texto.find(marcador_vinculo, inicio_vinculo + len(marcador_vinculo))
                if proximo_vinculo == -1:
                    fim_bloco = len(texto)
                else:
                    fim_bloco = proximo_vinculo
                
                bloco_vinculo = texto[inicio_vinculo:fim_bloco]
                
                # Extrair Seq e Código Emp deste bloco
                seq_match = re.search(r"Seq\.\s+.*?\n(\d+)", bloco_vinculo)
                codigo_match = re.search(r"Código Emp\.\s+.*?\n\d+\s+[0-9.\-]+\s+([0-9./\-]+)", bloco_vinculo)
                
                seq = seq_match.group(1) if seq_match else ""
                codigo_emp = codigo_match.group(1) if codigo_match else ""
                
                # Procurar seção de remunerações neste bloco
                inicio_remun = bloco_vinculo.find(marcador_remuneracao)
                if inicio_remun != -1:
                    secao_remun = bloco_vinculo[inicio_remun:]
                    
                    # Extrair linhas de remuneração
                    # Formato esperado: 
                    # Competência Remuneração Indicadores
                    # 01/1995 286,25 (vazio ou texto)
                    # 02/1995 286,25 (vazio ou texto)
                    
                    linhas = secao_remun.split('\n')
                    
                    # Pular cabeçalho (primeira linha com "Remunerações" e linha de títulos)
                    i = 0
                    while i < len(linhas) and not re.search(r'\d{2}/\d{4}', linhas[i]):
                        i += 1
                    
                    # Processar linhas de dados
                    while i < len(linhas):
                        linha = linhas[i].strip()
                        
                        # Parar se encontrar início de novo bloco ou seção
                        if not linha or "Matrícula" in linha or "Vínculos" in linha:
                            break
                        
                        # Tentar extrair: competência (mm/aaaa), valor, indicadores
                        # Padrão: "01/1995 286,25" ou "01/1995 286,25 texto_indicador"
                        match = re.match(r'(\d{2}/\d{4})\s+([\d.,]+)\s*(.*)', linha)
                        
                        if match:
                            competencia = match.group(1)
                            remuneracao = match.group(2).replace('.', '').replace(',', '.')  # Converte formato BR para numérico
                            indicadores = match.group(3).strip()
                            
                            registros_remuneracao.append({
                                "pagina": pagina_idx,
                                "seq": seq,
                                "codigo_emp": codigo_emp,
                                "competencia": competencia,
                                "remuneracao": remuneracao,
                                "indicadores": indicadores,
                            })
                        
                        i += 1
                
                pos = fim_bloco
    
    return registros_remuneracao


def extrair_remuneracoes_tabelas(linhas_saida) -> list[dict]:
    """Extrai remunerações das tabelas extraídas do PDF.
    
    Complementa a extração por texto, processando as tabelas estruturadas.
    """
    
    registros = []
    seq_atual = ""
    codigo_emp_atual = ""
    
    for linha in linhas_saida:
        pagina_idx, tabela_idx, *cols = linha
        
        # Detectar se é linha de vínculo (para pegar seq e código)
        texto_linha = " ".join(str(c) for c in cols)
        
        if "Matrícula do Tipo Filiado" in texto_linha:
            # Tentar extrair seq e código emp
            seq_match = re.search(r'\b(\d+)\b', texto_linha)
            codigo_match = re.search(r'([0-9]{2}\.[0-9]{3}\.[0-9]{3}/[0-9]{4}-[0-9]{2})', texto_linha)
            
            if seq_match:
                seq_atual = seq_match.group(1)
            if codigo_match:
                codigo_emp_atual = codigo_match.group(1)
        
        # Detectar linhas de remuneração: devem ter competência (mm/aaaa) e valor
        for col in cols:
            col_str = str(col).strip()
            if re.match(r'\d{2}/\d{4}', col_str):
                # Possível linha de remuneração
                # Tentar extrair competência e valor da mesma célula ou células adjacentes
                match = re.search(r'(\d{2}/\d{4})\s+([\d.,]+)', col_str)
                
                if match:
                    competencia = match.group(1)
                    remuneracao = match.group(2).replace('.', '').replace(',', '.')
                    
                    # Procurar indicadores nas células seguintes
                    idx = cols.index(col)
                    indicadores = ""
                    if idx + 1 < len(cols):
                        indicadores = str(cols[idx + 1]).strip()
                    
                    registros.append({
                        "pagina": pagina_idx,
                        "seq": seq_atual,
                        "codigo_emp": codigo_emp_atual,
                        "competencia": competencia,
                        "remuneracao": remuneracao,
                        "indicadores": indicadores,
                    })
    
    return registros


def salvar_remuneracoes_csv(caminho_pdf: str, linhas_saida, caminho_csv: str) -> None:
    """Gera um CSV com todas as remunerações extraídas por vínculo."""
    
    csv_path = Path(caminho_csv)
    
    # Extrair remunerações por texto (método principal)
    remuneracoes_texto = extrair_remuneracoes_texto(caminho_pdf)
    
    # Extrair remunerações das tabelas (complementar)
    remuneracoes_tabelas = extrair_remuneracoes_tabelas(linhas_saida)
    
    # Combinar e remover duplicatas
    todas_remuneracoes = remuneracoes_texto + remuneracoes_tabelas
    
    # Deduplica por (seq, codigo_emp, competencia)
    vistos = set()
    unicas = []
    
    for r in todas_remuneracoes:
        chave = (r.get("seq", ""), r.get("codigo_emp", ""), r.get("competencia", ""))
        if chave not in vistos and chave[2]:  # Ignora se não tem competência
            vistos.add(chave)
            unicas.append(r)
    
    # Ordenar por seq, codigo_emp e competencia
    unicas.sort(key=lambda x: (x.get("seq", ""), x.get("codigo_emp", ""), x.get("competencia", "")))
    
    # Salvar CSV
    with csv_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["Pagina", "Seq", "CodigoEmp", "Competencia", "Remuneracao", "Indicadores"])
        
        for r in unicas:
            writer.writerow([
                r.get("pagina", ""),
                r.get("seq", ""),
                r.get("codigo_emp", ""),
                r.get("competencia", ""),
                r.get("remuneracao", ""),
                r.get("indicadores", ""),
            ])


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

    # Dados do cabeçalho (Identificação do Filiado)
    cabecalho_path = str(pasta / f"{base}_dados_cliente.csv")
    dados_cab = extrair_dados_cabecalho(pdf_in)
    salvar_cabecalho_csv(dados_cab, cabecalho_path)

    # CSV adicional com os blocos de vínculos (texto bruto)
    vinculos_brutos_path = str(pasta / f"{base}_vinculos_brutos.csv")
    salvar_vinculos_brutos(linhas_saida, vinculos_brutos_path)

    # CSV estruturado com um vínculo por linha (inclui dados do cliente, se houver)
    vinculos_struct_path = str(pasta / f"{base}_vinculos_estruturado.csv")
    salvar_vinculos_estruturados(linhas_saida, vinculos_struct_path, dados_cab, pdf_in)

    # CSV com as remunerações de cada vínculo
    remuneracoes_path = str(pasta / f"{base}_remuneracoes.csv")
    salvar_remuneracoes_csv(pdf_in, linhas_saida, remuneracoes_path)

    print(f"Arquivo CSV bruto gerado em: {csv_out}")
    print(f"Arquivo de vínculos (texto bruto) gerado em: {vinculos_brutos_path}")
    print(f"Arquivo de vínculos estruturados gerado em: {vinculos_struct_path}")
    print(f"Arquivo de dados do cliente (cabeçalho) gerado em: {cabecalho_path}")
    print(f"Arquivo de remunerações gerado em: {remuneracoes_path}")


if __name__ == "__main__":
    main()

