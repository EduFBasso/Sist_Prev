"""
Extrator CLT - Remunerações (3 campos)

Processa seção "Remunerações" de vínculos com empregador (CLT).

Formato: Competência | Remuneração | Indicadores
- Até 3 competências por linha
- codigo_emp = CNPJ da empresa
"""

import re
from .utils import limpar_remuneracao, validar_valor


def processar_remuneracoes(secao_remun: str, seq: str, codigo_emp: str, 
                           pagina_idx: int, registros: list) -> None:
    """Processa seção de REMUNERAÇÕES (vínculos CLT com empregador).
    
    FORMATO CLT (3 campos):
        Remunerações
        Competência  Remuneração  Indicadores
        01/1995      286,25       13º SALÁRIO
        02/1995      286,25
        03/1995      286,25       MÚLTIPLOS VÍNCULOS
    
    REGEX PATTERN:
        (\\d{2}/\\d{4})          # Competência (captura)
        \\s+([\\d.,]+)           # Remuneração (captura)
        \\s*([^\\d/]*?)          # Indicadores (captura)
        (?=\\d{2}/\\d{4}|$)      # Lookahead próxima competência ou fim
    
    CARACTERÍSTICAS:
        - 3 campos por competência
        - Até 3 competências por linha
        - codigo_emp = CNPJ da empresa
    
    Args:
        secao_remun: Texto da seção "Remunerações" até próximo vínculo
        seq: Número de sequência do vínculo (ex: "2")
        codigo_emp: CNPJ da empresa (ex: "59.772.269/0001-39")
        pagina_idx: Número da página do PDF (1-indexed)
        registros: Lista onde adicionar dicionários de remunerações
    """
    linhas = secao_remun.split('\\n')
    
    # Pular cabeçalho até encontrar primeira competência
    i = 0
    while i < len(linhas) and not re.search(r'\\d{2}/\\d{4}', linhas[i]):
        i += 1
    
    # Processar linhas de dados
    while i < len(linhas):
        linha = linhas[i].strip()
        
        # Parar se encontrar marcador de novo bloco
        if not linha or "Matrícula" in linha or "Vínculos" in linha:
            break
        
        # Regex 3 campos: Competência + Remuneração + Indicadores
        # Até 3 competências por linha
        padroes = re.findall(
            r'(\\d{2}/\\d{4})\\s+([\\d.,]+)\\s*([^\\d/]*?)(?=\\d{2}/\\d{4}|$)',
            linha
        )
        
        if padroes:
            for competencia, remuneracao, indicadores in padroes:
                remuneracao_limpa = limpar_remuneracao(remuneracao)
                indicadores_limpo = indicadores.strip()
                
                try:
                    valor = float(remuneracao_limpa)
                    if validar_valor(valor):
                        registros.append({
                            "pagina": pagina_idx,
                            "seq": seq,
                            "codigo_emp": codigo_emp,  # CNPJ da empresa
                            "competencia": competencia,
                            "remuneracao": remuneracao_limpa,
                            "indicadores": indicadores_limpo,
                        })
                except ValueError:
                    # Valor inválido, ignorar
                    pass
        
        i += 1


def processar_valores_soltos_clt(zona_util: list[str], seq: str, codigo_emp: str,
                                  pagina_idx: int, registros: list) -> None:
    """Processa valores soltos no topo da zona útil (continuação de página anterior).
    
    Quando um vínculo CLT continua na próxima página, os valores aparecem soltos
    no topo antes do marcador "Matrícula do Tipo Filiado".
    
    Args:
        zona_util: Linhas da zona útil da página
        seq: Seq do vínculo da página anterior
        codigo_emp: CNPJ do vínculo da página anterior
        pagina_idx: Número da página
        registros: Lista de registros
    """
    for linha in zona_util:
        # Parar ao encontrar novo vínculo
        if "Matrícula do Tipo Filiado" in linha:
            break
        
        # Regex 3 campos (mesmo padrão CLT)
        padroes = re.findall(
            r'(\\d{2}/\\d{4})\\s+([\\d.,]+)\\s*([^\\d/]*?)(?=\\d{2}/\\d{4}|$)',
            linha
        )
        
        if padroes:
            for competencia, remuneracao, indicadores in padroes:
                remuneracao_limpa = limpar_remuneracao(remuneracao)
                
                try:
                    valor = float(remuneracao_limpa)
                    if validar_valor(valor):
                        registros.append({
                            "pagina": pagina_idx,
                            "seq": seq,
                            "codigo_emp": codigo_emp,
                            "competencia": competencia,
                            "remuneracao": remuneracao_limpa,
                            "indicadores": indicadores.strip(),
                        })
                except ValueError:
                    pass
