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
        2. Extrai blocos de vínculos
        3. Detecta tipo de cada bloco
        4. Delega para processador específico:
           - CLT → processar_remuneracoes_clt()
           - FACULTATIVO → processar_contribuicoes_facultativo()
           - Outros → log e pula
        5. Consolida todos os registros
    
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
        178 remunerações (163 CLT + 15 Facultativo)
    """
    registros: List[Dict[str, Any]] = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for pagina_idx, pagina in enumerate(pdf.pages, start=1):
            texto_pagina = pagina.extract_text() or ""
            
            # Processar blocos de vínculos nesta página
            blocos_vinculos = _extrair_blocos_vinculos(texto_pagina)
            
            for bloco in blocos_vinculos:
                # Detectar tipo do bloco
                tipo = detectar_tipo_vinculo(bloco['texto'])
                nome_tipo = obter_nome_tipo(tipo)
                
                # Delegar para processador específico
                if tipo == TipoVinculo.CLT:
                    _processar_bloco_clt(
                        bloco=bloco,
                        pagina=pagina_idx,
                        registros=registros
                    )
                
                elif tipo == TipoVinculo.FACULTATIVO:
                    _processar_bloco_facultativo(
                        bloco=bloco,
                        pagina=pagina_idx,
                        registros=registros
                    )
                
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


def _processar_bloco_clt(bloco: Dict[str, Any], pagina: int, registros: List[Dict[str, Any]]):
    """
    Processa bloco CLT extraindo remunerações.
    
    DELEGAÇÃO:
        Usa processar_remuneracoes_clt() de tipos/clt.py
        
    SEÇÃO PROCESSADA:
        "Remunerações" com 3 campos:
        MM/AAAA | Remuneração | Indicadores
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
