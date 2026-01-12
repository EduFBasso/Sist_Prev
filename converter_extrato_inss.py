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

    with csv_path.open("w", newline="", encoding="cp1252", errors="replace") as f:
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

    with csv_path.open("w", newline="", encoding="cp1252", errors="replace") as f:
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

    Captura TODOS os tipos de vínculos:
    - Vínculos com empresa: "Matrícula do Tipo Filiado no"
    - Vínculos facultativos: "Seq. NIT Origem do Vínculo Tipo Filiado no Vínculo"
    
    Retorna uma lista de tuplas (pagina_idx, texto_bloco).
    """

    pdf_path = Path(caminho_pdf)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF não encontrado: {pdf_path}")

    blocos: list[tuple[int, str]] = []

    with pdfplumber.open(pdf_path) as pdf:
        for pagina_idx, pagina in enumerate(pdf.pages, start=1):
            texto = pagina.extract_text() or ""
            
            # TIPO 1: Vínculos com empresa (formato antigo)
            marcador1 = "Matrícula do Tipo Filiado no"
            pos = 0
            while True:
                inicio = texto.find(marcador1, pos)
                if inicio == -1:
                    break

                # Bloco termina antes de "Remunerações" ou próximo vínculo
                fim = texto.find("Remunerações", inicio)
                if fim == -1:
                    # Procura próximo vínculo
                    proximo = texto.find(marcador1, inicio + len(marcador1))
                    if proximo != -1:
                        fim = proximo
                    else:
                        fim = len(texto)

                bloco = texto[inicio:fim]
                blocos.append((pagina_idx, bloco))
                pos = fim

            # TIPO 2: Vínculos facultativos (sem empresa)
            marcador2 = "Seq. NIT Origem do Vínculo Tipo Filiado no Vínculo"
            pos = 0
            while True:
                inicio = texto.find(marcador2, pos)
                if inicio == -1:
                    break

                # Captura as próximas linhas até "Contribuições" ou próximo vínculo
                fim = texto.find("Contribuições", inicio)
                if fim == -1:
                    # Procura próximo vínculo
                    proximo = texto.find("Seq. NIT", inicio + len(marcador2))
                    if proximo != -1:
                        fim = proximo
                    else:
                        # Busca por indicadores do fim (avisos INSS)
                        fim_alt = texto.find("O INSS poderá rever", inicio)
                        fim = fim_alt if fim_alt != -1 else len(texto)

                bloco = texto[inicio:fim]
                blocos.append((pagina_idx, bloco))
                pos = fim

    return blocos


def parse_vinculo_texto(texto: str) -> dict:
    """Extrai campos estruturados de QUALQUER tipo de vínculo do CNIS.

    Funciona para todos os tipos:
    - Vínculos com empresa (Seq. 1-10): "Empregado ou Agente Público"
    - Vínculos facultativos (Seq. 11-12): Sem empresa, com indicadores
    - Pré-facultativo concedido (Seq. 13+): PRE-FACULTCONC

    Retorna um dicionário com:
      seq, nit, codigo_emp, empresa, tipo_filiado, data_inicio, data_fim, 
      ult_remun, indicadores
    """

    try:
        linhas = str(texto).splitlines()
        if len(linhas) < 1:
            return {}
        
        # Normaliza o texto completo - junta TODAS as linhas primeiro
        corpo_completo = " ".join(linhas).strip()
        corpo_completo = " ".join(corpo_completo.split())
        
        # Para vínculos facultativos, o formato é mais direto (sem pular linhas)
        # Detecta se é formato facultativo ANTES de processar
        if "Seq. NIT Origem do Vínculo" in corpo_completo:
            # FORMATO FACULTATIVO - usa todas as linhas
            corpo = corpo_completo
        else:
            # FORMATO COM EMPRESA - pula cabeçalho (primeiras 2 linhas)
            corpo = " ".join(linhas[2:]).strip() if len(linhas) > 2 else corpo_completo
            corpo = " ".join(corpo.split())
        
        # Identificação de datas completas (dd/mm/aaaa) e competência (mm/aaaa)
        def is_data_completa(token: str) -> bool:
            return re.fullmatch(r"\d{2}/\d{2}/\d{4}", token) is not None

        def is_competencia(token: str) -> bool:
            return re.fullmatch(r"\d{2}/\d{4}", token) is not None
        
        # Inicializa campos
        seq = ""
        nit = ""
        codigo_emp = ""
        empresa = ""
        tipo_filiado = ""
        data_inicio = ""
        data_fim = ""
        ult_remun = ""
        indicadores = ""
        
        # DETECTA FORMATO DO VÍNCULO
        
        # FORMATO FACULTATIVO: "11 125.37781.66-1 RECOLHIMENTO Facultativo 01/09/2019 31/10/2019 IREC-INDPEND"
        # Padrão: Seq NIT ORIGEM TipoFiliado DataInicio DataFim Indicadores
        if "RECOLHIMENTO" in corpo and "Facultativo" in corpo:
            # Extrai usando padrão direto
            # Formato: Seq NIT RECOLHIMENTO Facultativo dd/mm/aaaa dd/mm/aaaa INDICADOR
            match = re.search(
                r'(\d+)\s+(\d{3}\.\d{5}\.\d{2}-\d)\s+RECOLHIMENTO\s+Facultativo\s+(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})\s*(\S+)?',
                corpo
            )
            if match:
                seq = match.group(1)
                nit = match.group(2)
                tipo_filiado = "Contribuinte Facultativo"
                data_inicio = match.group(3)
                data_fim = match.group(4)
                indicadores = match.group(5) if match.group(5) else ""
                
                return {
                    "seq": seq,
                    "nit": nit,
                    "codigo_emp": "",
                    "empresa": "",
                    "tipo_filiado": tipo_filiado,
                    "data_inicio": data_inicio,
                    "data_fim": data_fim,
                    "ult_remun": "",
                    "indicadores": indicadores,
                }
        
        # FORMATO COM EMPRESA (original)
        # 1) EXTRAÇÃO DA SEQUÊNCIA (primeiro número depois do cabeçalho)
        match_seq = re.search(r'\b(\d+)\s+(\d{3}\.\d{5}\.\d{2}-\d)', corpo)
        if match_seq:
            seq = match_seq.group(1)
            nit = match_seq.group(2)
        else:
            # Formato alternativo sem NIT visível
            match_seq = re.search(r'^(\d+)\s+', corpo)
            if match_seq:
                seq = match_seq.group(1)
        
        # 2) IDENTIFICAÇÃO DO TIPO DE VÍNCULO
        
        # Tipo 1: Vínculo com empresa (CNPJ presente)
        cnpj_pattern = r'(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})'
        match_cnpj = re.search(cnpj_pattern, corpo)
        
        if match_cnpj:
            # VÍNCULO COM EMPRESA
            codigo_emp = match_cnpj.group(1)
            
            # Detecta tipo de filiado
            if "Empregado ou Agente" in corpo:
                tipo_filiado = "Empregado ou Agente Público"
                
                # Extrai empresa (texto entre CNPJ e tipo de filiado)
                antes_tipo = corpo.split("Empregado ou Agente", 1)[0]
                depois_cnpj = antes_tipo.split(codigo_emp, 1)[1].strip() if codigo_emp in antes_tipo else ""
                empresa = " ".join(depois_cnpj.split())
            
            elif "Contribuinte Individual" in corpo:
                tipo_filiado = "Contribuinte Individual"
                antes_tipo = corpo.split("Contribuinte Individual", 1)[0]
                depois_cnpj = antes_tipo.split(codigo_emp, 1)[1].strip() if codigo_emp in antes_tipo else ""
                empresa = " ".join(depois_cnpj.split())
        
        else:
            # VÍNCULO SEM EMPRESA (Facultativo, PRE-FACULTCONC, etc.)
            
            # Detecta indicadores especiais
            if "PRE-FACULTCONC" in corpo or "PREC-FACULTCONC" in corpo:
                tipo_filiado = "Pré-Facultativo Concedido"
                # Extrai as competências do indicador
                match_ind = re.findall(r'(?:PRE|PREC)-FACULTCONC', corpo)
                if match_ind:
                    indicadores = "PRE-FACULTCONC"
            
            elif "Facultativo" in corpo or "FACULTATIVO" in corpo:
                tipo_filiado = "Contribuinte Facultativo"
            
            elif "Contribuições" in corpo or "Recolhimento" in corpo:
                tipo_filiado = "Recolhimento Facultativo"
            
            else:
                # Tipo genérico - tenta extrair do texto
                tipo_match = re.search(r'((?:[A-Z][a-z]+\s*){2,4})\s+\d{2}/\d{2}/\d{4}', corpo)
                if tipo_match:
                    tipo_filiado = tipo_match.group(1).strip()
        
        # 3) EXTRAÇÃO DE DATAS (dd/mm/aaaa) e COMPETÊNCIAS (mm/aaaa)
        tokens = corpo.split()
        datas = [t for t in tokens if is_data_completa(t) or is_competencia(t)]
        
        if datas:
            # Filtra apenas datas completas primeiro (dd/mm/aaaa)
            datas_completas = [d for d in datas if is_data_completa(d)]
            competencias = [d for d in datas if is_competencia(d)]
            
            if len(datas_completas) >= 2:
                # Padrão típico: Data Início, Data Fim
                data_inicio = datas_completas[0]
                data_fim = datas_completas[1]
            elif len(datas_completas) == 1:
                # Apenas uma data completa
                data_inicio = datas_completas[0]
            
            # Última remuneração é a última competência (mm/aaaa)
            if competencias:
                ult_remun = competencias[-1]
        
        # Se não encontrou datas completas, mas tem seq
        if not datas and seq:
            return {
                "seq": seq,
                "nit": nit,
                "codigo_emp": codigo_emp,
                "empresa": empresa,
                "tipo_filiado": tipo_filiado or "Desconhecido",
                "data_inicio": data_inicio,
                "data_fim": data_fim,
                "ult_remun": ult_remun,
                "indicadores": indicadores,
            }
        
        # Valida se pelo menos tem sequência E alguma data
        if not seq or not datas:
            return {}
        
        return {
            "seq": seq,
            "nit": nit,
            "codigo_emp": codigo_emp,
            "empresa": empresa,
            "tipo_filiado": tipo_filiado,
            "data_inicio": data_inicio,
            "data_fim": data_fim,
            "ult_remun": ult_remun,
            "indicadores": indicadores,
        }
    
    except Exception as e:
        # Em caso de erro, retorna vazio
        return {}


def salvar_vinculos_estruturados(linhas_saida, caminho_csv: str, dados_cab: dict | None = None, caminho_pdf: str | None = None) -> None:
    """Gera um CSV estruturado com TODOS os vínculos (todas as sequências).

    Captura vínculos com empresa, facultativos, pré-facultativos, etc.
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

    # 2) Complementar com vínculos extraídos diretamente do texto do PDF (PRINCIPAL)
    if caminho_pdf:
        for pagina_idx, bloco in extrair_vinculos_texto(caminho_pdf):
            dados = parse_vinculo_texto(bloco)
            if dados:
                registros.append({
                    "pagina": pagina_idx,
                    "tabela": 0,
                    **dados,
                })

    # 3) Remover duplicidades
    #    Chave de deduplicação: (Seq, NIT, CodigoEmp, DataInicio)
    #    Usa DataInicio para diferenciar múltiplos vínculos sem CNPJ
    vistos: set[tuple[str, str, str, str]] = set()
    unicos: list[dict] = []

    for r in registros:
        chave = (
            str(r.get("seq", "")), 
            str(r.get("nit", "")), 
            str(r.get("codigo_emp", "")),
            str(r.get("data_inicio", ""))
        )
        if chave in vistos:
            continue
        vistos.add(chave)
        unicos.append(r)

    # Ordena por sequência (numérica)
    unicos.sort(key=lambda x: int(x.get("seq", "0")) if x.get("seq", "").isdigit() else 999)

    # 4) Gravar CSV final estruturado (cp1252 para Windows/VBA)
    with csv_path.open("w", newline="", encoding="cp1252", errors="replace") as f:
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
            "Indicadores",
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
                r.get("indicadores", ""),
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

    with csv_path.open("w", newline="", encoding="cp1252", errors="replace") as f:
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
    
    # Salvar CSV (cp1252 para Windows/VBA)
    with csv_path.open("w", newline="", encoding="cp1252", errors="replace") as f:
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


def gerar_relatorio_validacao(caminho_pdf: str, caminho_saida: str) -> dict:
    """Gera relatório de validação da extração.
    
    Retorna estatísticas sobre:
    - Total de blocos encontrados
    - Total de vínculos parseados
    - Sequências capturadas
    - Blocos não parseados (para análise)
    - Tipos de vínculo únicos
    """
    
    blocos = extrair_vinculos_texto(caminho_pdf)
    
    vinculos_parseados = []
    blocos_nao_parseados = []
    sequencias_encontradas = set()
    tipos_vinculo = set()
    
    for pag, bloco in blocos:
        dados = parse_vinculo_texto(bloco)
        
        if dados and dados.get("seq"):
            vinculos_parseados.append(dados)
            sequencias_encontradas.add(dados["seq"])
            if dados.get("tipo_filiado"):
                tipos_vinculo.add(dados["tipo_filiado"])
        else:
            # Preview das primeiras 200 caracteres do bloco não parseado
            preview = " ".join(bloco.split()[:30])
            blocos_nao_parseados.append({
                "pagina": pag,
                "preview": preview
            })
    
    # Ordena sequências numericamente
    sequencias_ordenadas = sorted(
        [int(s) for s in sequencias_encontradas if s.isdigit()]
    )
    
    relatorio = {
        "total_blocos": len(blocos),
        "total_parseados": len(vinculos_parseados),
        "total_nao_parseados": len(blocos_nao_parseados),
        "sequencias": sequencias_ordenadas,
        "tipos_vinculo": sorted(tipos_vinculo),
        "blocos_nao_parseados": blocos_nao_parseados,
    }
    
    # Salva relatório em arquivo de texto (cp1252 para Windows)
    with open(caminho_saida, "w", encoding="cp1252", errors="replace") as f:
        f.write("=" * 80 + "\n")
        f.write("RELATÓRIO DE VALIDAÇÃO DA EXTRAÇÃO - CNIS\n")
        f.write("=" * 80 + "\n\n")
        
        f.write(f"📊 RESUMO\n")
        f.write(f"  • Total de blocos encontrados: {relatorio['total_blocos']}\n")
        f.write(f"  • Vínculos parseados com sucesso: {relatorio['total_parseados']}\n")
        f.write(f"  • Blocos NÃO parseados: {relatorio['total_nao_parseados']}\n\n")
        
        f.write(f"📋 SEQUÊNCIAS CAPTURADAS ({len(sequencias_ordenadas)})\n")
        if sequencias_ordenadas:
            f.write(f"  {', '.join(map(str, sequencias_ordenadas))}\n\n")
        else:
            f.write("  (nenhuma sequência identificada)\n\n")
        
        f.write(f"🏷️  TIPOS DE VÍNCULO ENCONTRADOS ({len(relatorio['tipos_vinculo'])})\n")
        for tipo in relatorio['tipos_vinculo']:
            f.write(f"  • {tipo}\n")
        f.write("\n")
        
        if blocos_nao_parseados:
            f.write(f"⚠️  BLOCOS NÃO PARSEADOS ({len(blocos_nao_parseados)})\n")
            f.write("   (Analise estes blocos para identificar novos formatos)\n\n")
            
            for i, bloco in enumerate(blocos_nao_parseados, 1):
                f.write(f"  [{i}] Página {bloco['pagina']}:\n")
                f.write(f"      {bloco['preview']}...\n\n")
        else:
            f.write("✅ TODOS OS BLOCOS FORAM PARSEADOS COM SUCESSO!\n\n")
        
        f.write("=" * 80 + "\n")
        f.write("💡 DICA: Se há blocos não parseados, envie o PDF para análise.\n")
        f.write("   O extrator pode precisar de ajustes para novos formatos.\n")
        f.write("=" * 80 + "\n")
    
    return relatorio


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

    # Relatório de validação
    validacao_path = str(pasta / f"{base}_validacao.txt")
    relatorio = gerar_relatorio_validacao(pdf_in, validacao_path)

    print(f"Arquivo CSV bruto gerado em: {csv_out}")
    print(f"Arquivo de vínculos (texto bruto) gerado em: {vinculos_brutos_path}")
    print(f"Arquivo de vínculos estruturados gerado em: {vinculos_struct_path}")
    print(f"Arquivo de dados do cliente (cabeçalho) gerado em: {cabecalho_path}")
    print(f"Arquivo de remunerações gerado em: {remuneracoes_path}")
    print(f"Relatório de validação gerado em: {validacao_path}")
    print()
    print(f"✅ Extração concluída:")
    print(f"   • {relatorio['total_parseados']} vínculos capturados")
    print(f"   • Sequências: {', '.join(map(str, relatorio['sequencias']))}")
    if relatorio['total_nao_parseados'] > 0:
        print(f"   ⚠️  {relatorio['total_nao_parseados']} blocos não parseados - veja {validacao_path}")


if __name__ == "__main__":
    main()

