"""
Coordenador de Remunerações - Delegação Automática por Tipo

RESPONSABILIDADE:
    Orquestrar extração de remunerações delegando automaticamente
    para o processador correto baseado no tipo de vínculo detectado.

ESTRATÉGIA:
    1. Itera páginas do PDF
    2. Extrai blocos de vínculos
    3. Detecta tipo de cada bloco (CLT, FACULTATIVO, etc)
    4. Delega para processador especializado
    5. Consolida resultados

TIPOS SUPORTADOS NO MOMENTO:
    ✅ CLT - processar_remuneracoes_clt() - 3 campos
    ✅ FACULTATIVO - processar_contribuicoes_facultativo() - 5 campos
    
EXPANSÃO FUTURA:
    Para adicionar novo tipo:
    1. Criar processador em tipos/novo_tipo.py
    2. Adicionar em TipoVinculo enum
    3. Adicionar case nesta função
    4. Testar com documento real

SUBSTITUIÇÃO:
    Este módulo substitui extrair_remuneracoes_texto() monolítico (560 linhas)
    do converter_extrato_inss.py, mantendo baseline: 178 remunerações.
"""

import re
import pdfplumber
from typing import List, Dict, Any

from .detector import detectar_tipo_vinculo, TipoVinculo, obter_nome_tipo
from .clt import processar_remuneracoes_clt, processar_valores_soltos_clt
from .facultativo import processar_contribuicoes_facultativo


def extrair_remuneracoes_coordenado(pdf_path: str) -> List[Dict[str, Any]]:
    """
    Extrai todas as remunerações do PDF delegando automaticamente
    para o processador correto baseado no tipo de vínculo.
    
    ALGORITMO:
        1. Abre PDF e itera páginas
        2. Extrai ZONA ÚTIL (entre cabeçalho e rodapé)
        3. Processa valores soltos no topo (continuação página anterior)
        4. Extrai blocos de vínculos
        5. Detecta tipo de cada bloco
        6. Delega para processador específico:
           - CLT → processar_remuneracoes_clt()
           - FACULTATIVO → processar_contribuicoes_facultativo()
           - Outros → log e pula
        7. Consolida todos os registros
    
    Args:
        pdf_path: Caminho do arquivo PDF do CNIS
        
    Returns:
        List[Dict]: Lista de remunerações extraídas
        Formato: [
            {
                'seq': '1',
                'cnpj': '12.345.678/0001-90',
                'competencia': '01/2020',
                'remuneracao': '2500.00',
                'indicadores': 'texto',
                'pagina': 1,
                'tipo_vinculo': 'CLT'
            },
            ...
        ]
        
    BASELINE ESPERADO:
        178 remunerações (todos CLT)
    """
    registros: List[Dict[str, Any]] = []
    ultimo_seq_pagina_anterior = None
    ultimo_cnpj_pagina_anterior = None
    bloco_cortado_pagina_anterior = None  # Bloco que começou mas foi cortado pelo rodapé
    
    with pdfplumber.open(pdf_path) as pdf:
        for pagina_idx, pagina in enumerate(pdf.pages, start=1):
            texto_pagina = pagina.extract_text() or ""
            
            # Extrair zona útil (entre cabeçalho e rodapé)
            zona_util = _extrair_zona_util(texto_pagina)
            
            # Se tinha bloco cortado na página anterior, processar sua seção "Remunerações" aqui
            if bloco_cortado_pagina_anterior:
                _processar_secao_cortada(
                    zona_util=zona_util,
                    bloco_cortado=bloco_cortado_pagina_anterior,
                    pagina=pagina_idx,
                    registros=registros
                )
                bloco_cortado_pagina_anterior = None
            
            # Processar valores soltos no topo (continuação de página anterior)
            elif pagina_idx > 1 and ultimo_seq_pagina_anterior:
                _processar_valores_soltos_continuacao(
                    zona_util=zona_util,
                    seq=ultimo_seq_pagina_anterior,
                    cnpj=ultimo_cnpj_pagina_anterior,
                    pagina=pagina_idx,
                    registros=registros
                )
            
            # Processar blocos de vínculos nesta página
            blocos_vinculos = _extrair_blocos_vinculos(texto_pagina)
            
            for bloco in blocos_vinculos:
                # Detectar tipo do bloco
                tipo = detectar_tipo_vinculo(bloco['texto'])
                nome_tipo = obter_nome_tipo(tipo)
                
                # Delegar para processador específico
                if tipo == TipoVinculo.CLT:
                    tem_remuneracoes = _processar_bloco_clt(
                        bloco=bloco,
                        pagina=pagina_idx,
                        registros=registros
                    )
                    
                    # Se bloco não tem seção "Remunerações", foi cortado pelo rodapé
                    if not tem_remuneracoes:
                        bloco_cortado_pagina_anterior = bloco
                    else:
                        # Guardar para próxima página (valores soltos)
                        ultimo_seq_pagina_anterior = bloco['seq']
                        ultimo_cnpj_pagina_anterior = bloco['cnpj']
                
                elif tipo == TipoVinculo.FACULTATIVO:
                    _processar_bloco_facultativo(
                        bloco=bloco,
                        pagina=pagina_idx,
                        registros=registros
                    )
                    
                    # Guardar para próxima página
                    ultimo_seq_pagina_anterior = bloco['seq']
                    ultimo_cnpj_pagina_anterior = 'FACULTATIVO'
                
                elif tipo == TipoVinculo.DESCONHECIDO:
                    # Log para debug (não impede processamento)
                    print(f"⚠️  Bloco desconhecido na página {pagina_idx} (Seq {bloco.get('seq', '?')})")
                
                # Adicionar novos tipos aqui no futuro:
                # elif tipo == TipoVinculo.MEI:
                #     _processar_bloco_mei(bloco, pagina_idx, registros)
    
    return registros


def _extrair_blocos_vinculos(texto_pagina: str) -> List[Dict[str, Any]]:
    """
    Extrai blocos de vínculos de uma página de texto.
    
    MARCADOR DE INÍCIO:
        "Matrícula do Tipo Filiado no"
        
    ESTRUTURA DO BLOCO:
        Seq. NIT Código Emp. Origem do Vínculo...
        1 125.37781.66-1 12.345.678/0001-90 EMPRESA...
        ...
        Remunerações (ou Contribuições)
        01/2020 2500.00 Texto
        ...
        [próximo bloco ou fim]
    
    Returns:
        List[Dict]: Lista de dicionários com:
            - 'seq': Número de sequência
            - 'nit': NIT extraído
            - 'cnpj': CNPJ se presente (CLT)
            - 'texto': Texto completo do bloco
    """
    blocos = []
    marcador = "Matrícula do Tipo Filiado no"
    pos = 0
    
    while True:
        # Encontrar próximo marcador
        inicio = texto_pagina.find(marcador, pos)
        if inicio == -1:
            break
        
        # Encontrar próximo marcador (fim deste bloco)
        proximo = texto_pagina.find(marcador, inicio + len(marcador))
        if proximo == -1:
            # Último bloco da página
            bloco_texto = texto_pagina[inicio:]
        else:
            bloco_texto = texto_pagina[inicio:proximo]
        
        # Extrair metadados básicos do bloco
        seq = _extrair_seq_bloco(bloco_texto)
        nit = _extrair_nit_bloco(bloco_texto)
        cnpj = _extrair_cnpj_bloco(bloco_texto)
        
        blocos.append({
            'seq': seq,
            'nit': nit,
            'cnpj': cnpj,
            'texto': bloco_texto
        })
        
        pos = inicio + len(marcador)
    
    return blocos


def _extrair_seq_bloco(texto: str) -> str:
    """Extrai número de sequência do bloco."""
    # Procura por "Seq. NIT ..." e pega primeiro número da próxima linha
    match = re.search(r'Seq\.\s+NIT.*?\n\s*(\d+)\s+', texto, re.IGNORECASE)
    return match.group(1) if match else ""


def _extrair_nit_bloco(texto: str) -> str:
    """Extrai NIT do bloco."""
    # Formato: 123.45678.90-1
    match = re.search(r'\d{3}\.\d{5}\.\d{2}-\d', texto)
    return match.group(0) if match else ""


def _extrair_cnpj_bloco(texto: str) -> str:
    """Extrai CNPJ do bloco (se presente - indica CLT)."""
    # Formato: 12.345.678/0001-90
    match = re.search(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}', texto)
    return match.group(0) if match else ""


def _extrair_zona_util(texto_pagina: str) -> str:
    """
    Extrai a zona útil da página: entre cabeçalho e rodapé.
    
    ZONA ÚTIL:
        INÍCIO: Após "Identificação do Filiado" e dados (NIT, CPF, Nome, Data nasc, Nome mãe)
        FIM: Antes do rodapé "O INSS poderá rever a qualquer tempo..."
    
    OBJETIVO:
        Isolar apenas a área com dados de vínculos/remunerações,
        excluindo cabeçalho e rodapé que se repetem em todas as páginas.
    
    Args:
        texto_pagina: Texto completo da página
        
    Returns:
        str: Texto da zona útil
    """
    # Marcar início da zona útil (após cabeçalho)
    inicio = 0
    
    # Procurar fim do cabeçalho "Identificação do Filiado"
    match_cabecalho = re.search(r'Nome da mãe:', texto_pagina, re.IGNORECASE)
    if match_cabecalho:
        inicio = match_cabecalho.end()
    else:
        # Alternativa: procurar "Relações Previdenciárias"
        match_relacoes = re.search(r'Relações Previdenciárias', texto_pagina, re.IGNORECASE)
        if match_relacoes:
            inicio = match_relacoes.end()
    
    # Marcar fim da zona útil (antes do rodapé)
    fim = len(texto_pagina)
    
    # Procurar início do rodapé
    match_rodape = re.search(r'O INSS poderá rever a qualquer tempo', texto_pagina, re.IGNORECASE)
    if match_rodape:
        fim = match_rodape.start()
    
    return texto_pagina[inicio:fim]


def _processar_valores_soltos_continuacao(zona_util: str, seq: str, cnpj: str, 
                                          pagina: int, registros: List[Dict[str, Any]]):
    """
    Processa valores soltos no topo da página (continuação de página anterior).
    
    LÓGICA:
        Se trocou a página, e ao iniciar a zona útil identificamos MM/YYYY,
        e NÃO há cabeçalhos (Competência Remuneração Indicadores),
        então esses valores pertencem à última Seq da página anterior.
    
    GATILHOS:
        ✅ Tem competências (MM/YYYY)
        ❌ NÃO tem cabeçalho "Competência" ou "Remuneração"
        ❌ NÃO tem marcador de novo bloco "Matrícula do Tipo Filiado"
    
    Args:
        zona_util: Texto da zona útil da página
        seq: Seq da página anterior (último bloco processado)
        cnpj: CNPJ da página anterior
        pagina: Número da página atual
        registros: Lista de registros
    """
    # Extrair texto antes do primeiro bloco
    match_primeiro_bloco = re.search(r'Matrícula do Tipo Filiado', zona_util, re.IGNORECASE)
    
    if match_primeiro_bloco:
        # Pegar texto antes do primeiro bloco
        texto_topo = zona_util[:match_primeiro_bloco.start()]
    else:
        # Toda a zona útil (não há bloco nesta página)
        texto_topo = zona_util
    
    # Verificar se há competências no topo
    if not re.search(r'\d{2}/\d{4}', texto_topo):
        return  # Não há competências, nada a fazer
    
    # Verificar se NÃO é um bloco completo (sem cabeçalhos)
    if re.search(r'Competência\s+Remuneração', texto_topo, re.IGNORECASE):
        return  # Tem cabeçalho, não é continuação
    
    # Processar valores soltos (formato CLT - 3 campos)
    linhas = texto_topo.split('\n')
    
    for linha in linhas:
        linha = linha.strip()
        if not linha:
            continue
        
        # Regex 3 campos CLT: Competência + Remuneração + Indicadores
        padroes = re.findall(
            r'(\d{2}/\d{4})\s+([\d.,]+)\s*([^\d/]*?)(?=\d{2}/\d{4}|$)',
            linha
        )
        
        if padroes:
            for competencia, remuneracao, indicadores in padroes:
                from ..utils import limpar_remuneracao, validar_valor
                remuneracao_limpa = limpar_remuneracao(remuneracao)
                
                try:
                    valor = float(remuneracao_limpa)
                    if validar_valor(valor):
                        registros.append({
                            'pagina': pagina,
                            'seq': seq,
                            'codigo_emp': cnpj,
                            'cnpj': cnpj,
                            'competencia': competencia,
                            'remuneracao': remuneracao_limpa,
                            'indicadores': indicadores.strip(),
                            'tipo_vinculo': 'CLT'
                        })
                except ValueError:
                    pass


def _processar_secao_cortada(zona_util: str, bloco_cortado: Dict[str, Any], 
                             pagina: int, registros: List[Dict[str, Any]]):
    """
    Processa seção "Remunerações" que foi cortada pelo rodapé na página anterior.
    
    CENÁRIO:
        Página anterior: Cabeçalho do vínculo (Seq, NIT, CNPJ, Empresa, Datas) + RODAPÉ
        Página atual: Seção "Remunerações" com cabeçalho (Competência | Remuneração | Indicadores)
    
    LÓGICA:
        Se início da zona útil tem cabeçalho "Competência Remuneração Indicadores"
        (ou "Remunerações" seguido de cabeçalho),
        então processar como continuação do bloco cortado da página anterior.
    
    Args:
        zona_util: Texto da zona útil da página atual
        bloco_cortado: Bloco que começou na página anterior mas foi cortado
        pagina: Número da página atual
        registros: Lista de registros
    """
    seq = bloco_cortado['seq']
    cnpj = bloco_cortado['cnpj']
    
    # Verificar se início tem cabeçalho "Competência Remuneração" ou seção "Remunerações"
    tem_cabecalho_remun = re.search(
        r'(Remunerações\s*)?Competência\s+Remuneração',
        zona_util[:500],  # Primeiros 500 chars
        re.IGNORECASE
    )
    
    if not tem_cabecalho_remun:
        return  # Não é continuação de seção cortada
    
    # Extrair até o próximo bloco ou fim
    match_proximo_bloco = re.search(r'Matrícula do Tipo Filiado', zona_util, re.IGNORECASE)
    
    if match_proximo_bloco:
        secao_remun = zona_util[:match_proximo_bloco.start()]
    else:
        secao_remun = zona_util
    
    # Processar seção de remunerações
    idx_antes = len(registros)
    processar_remuneracoes_clt(secao_remun, seq, cnpj, pagina, registros)
    
    # Adicionar tipo_vinculo e cnpj
    for i in range(idx_antes, len(registros)):
        registros[i]['tipo_vinculo'] = 'CLT'
        registros[i]['cnpj'] = cnpj


def _processar_bloco_clt(bloco: Dict[str, Any], pagina: int, registros: List[Dict[str, Any]]) -> bool:
    """
    Processa bloco CLT extraindo remunerações.
    
    DELEGAÇÃO:
        Usa processar_remuneracoes_clt() de tipos/clt.py
        
    SEÇÃO PROCESSADA:
        "Remunerações" com 3 campos:
        MM/AAAA | Remuneração | Indicadores
        
    Returns:
        bool: True se encontrou seção "Remunerações", False se bloco cortado pelo rodapé
    """
    texto = bloco['texto']
    seq = bloco['seq']
    cnpj = bloco['cnpj']
    
    # Encontrar seção "Remunerações"
    match_remun = re.search(r'Remunerações\s*(.*?)(?:Contribuições|Matrícula|$)', 
                           texto, re.DOTALL | re.IGNORECASE)
    
    if match_remun:
        secao_remun = match_remun.group(1)
        # Delegar para processador CLT
        idx_antes = len(registros)
        processar_remuneracoes_clt(secao_remun, seq, cnpj, pagina, registros)
        
        # Adicionar tipo_vinculo e cnpj aos registros CLT
        for i in range(idx_antes, len(registros)):
            registros[i]['tipo_vinculo'] = 'CLT'
            registros[i]['cnpj'] = cnpj
        
        # Tentar capturar valores soltos também (converter texto para lista de linhas)
        idx_antes = len(registros)
        linhas_zona = texto.split('\n')
        processar_valores_soltos_clt(linhas_zona, seq, cnpj, pagina, registros)
        
        # Adicionar tipo_vinculo aos valores soltos
        for i in range(idx_antes, len(registros)):
            registros[i]['tipo_vinculo'] = 'CLT'
            registros[i]['cnpj'] = cnpj
        
        return True  # Seção encontrada
    else:
        # Bloco sem seção "Remunerações" = cortado pelo rodapé
        return False
        # Adicionar tipo_vinculo aos valores soltos
        for i in range(idx_antes, len(registros)):
            registros[i]['tipo_vinculo'] = 'CLT'
            registros[i]['cnpj'] = cnpj


def _processar_bloco_facultativo(bloco: Dict[str, Any], pagina: int, registros: List[Dict[str, Any]]):
    """
    Processa bloco FACULTATIVO extraindo contribuições.
    
    DELEGAÇÃO:
        Usa processar_contribuicoes_facultativo() de tipos/facultativo.py
        
    SEÇÃO PROCESSADA:
        "Contribuições" com 5 campos:
        MM/AAAA | Data Pagto | Salário | Contribuição | Indicadores
    """
    texto = bloco['texto']
    seq = bloco['seq']
    
    # Encontrar seção "Contribuições"
    match_contrib = re.search(r'Contribuições\s*(.*?)(?:Remunerações|Matrícula|$)', 
                             texto, re.DOTALL | re.IGNORECASE)
    
    if match_contrib:
        secao_contrib = match_contrib.group(1)
        # Delegar para processador Facultativo
        idx_antes = len(registros)
        processar_contribuicoes_facultativo(secao_contrib, seq, pagina, registros)
        
        # Adicionar tipo_vinculo aos registros Facultativo
        for i in range(idx_antes, len(registros)):
            registros[i]['tipo_vinculo'] = 'FACULTATIVO'
            registros[i]['cnpj'] = 'FACULTATIVO'  # Não tem CNPJ


# ========================================
# FUNÇÃO DE COMPATIBILIDADE COM CÓDIGO LEGADO
# ========================================

def extrair_remuneracoes_texto(pdf_path: str) -> List[Dict[str, Any]]:
    """
    Função de compatibilidade com código legado.
    
    OBJETIVO:
        Permitir substituição direta no converter_extrato_inss.py sem
        quebrar chamadas existentes.
        
    IMPLEMENTAÇÃO:
        Delega para extrair_remuneracoes_coordenado() que usa detecção
        automática de tipos.
    """
    return extrair_remuneracoes_coordenado(pdf_path)


# ========================================
# UTILITÁRIOS DE VALIDAÇÃO
# ========================================

def validar_baseline_remuneracoes(registros: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Valida se o número de remunerações está correto.
    
    BASELINE ESPERADO:
        - Total: 178 remunerações
        - CLT: 163 remunerações
        - FACULTATIVO: 15 remunerações
    
    Args:
        registros: Lista de remunerações extraídas
        
    Returns:
        Dict com estatísticas e validação:
        {
            'total': 178,
            'por_tipo': {'CLT': 163, 'FACULTATIVO': 15},
            'valido': True,
            'mensagem': '✅ Baseline preservado'
        }
    """
    total = len(registros)
    
    # Contar por tipo
    por_tipo = {}
    for reg in registros:
        tipo = reg.get('tipo_vinculo', 'DESCONHECIDO')
        por_tipo[tipo] = por_tipo.get(tipo, 0) + 1
    
    # Validar baseline
    baseline_ok = (
        total == 178 and
        por_tipo.get('CLT', 0) == 163 and
        por_tipo.get('FACULTATIVO', 0) == 15
    )
    
    return {
        'total': total,
        'por_tipo': por_tipo,
        'valido': baseline_ok,
        'mensagem': '✅ Baseline preservado (178 remunerações)' if baseline_ok 
                   else f'⚠️  Baseline diferente: {total} remunerações (esperado 178)'
    }
