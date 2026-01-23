"""
Extrator Facultativo - Contribuições (5 campos)

Processa seção "Contribuições" de vínculos Facultativos.

Formato: Competência | Data Pagto | Contribuição | Salário | Indicadores
- Até 2 competências por linha
- codigo_emp = "FACULTATIVO" (sem CNPJ)
- Contribuição → salva como remuneracao
"""

import re
from .utils import limpar_remuneracao, validar_valor


def processar_contribuicoes_facultativo(secao_contrib: str, seq: str, pagina_idx: int, 
                                       registros: list) -> None:
    """Processa seção de CONTRIBUIÇÕES (vínculos Facultativos).
    
    FORMATO FACULTATIVO (5 campos):
        Contribuições
        Competência   Data Pagto.   Contribuição   Salário Contrib.   Indicadores
        09/2019       15/09/2019    200,00         1045,00            PREC-FACULTCONC
        10/2019       15/10/2019    199,60         1045,00            PREC-FACULTCONC
    
    REGEX PATTERN:
        (\\d{2}/\\d{4})                      # Competência (captura)
        \\s+\\d{2}/\\d{2}/\\d{4}               # Data Pagto (ignora)
        \\s+([\\d.,]+)                       # Contribuição (captura como remuneracao)
        \\s+([\\d.,]+)                       # Salário (ignora)
        \\s*([^\\d/]*?)                      # Indicadores (captura)
        (?=\\d{2}/\\d{4}|$)                  # Lookahead próxima competência ou fim
    
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
    linhas = secao_contrib.split('\\n')
    
    # Pular cabeçalho até encontrar primeira competência
    i = 0
    while i < len(linhas) and not re.search(r'\\d{2}/\\d{4}', linhas[i]):
        i += 1
    
    # Processar linhas de dados
    while i < len(linhas):
        linha = linhas[i].strip()
        
        # Parar se encontrar marcador de novo vínculo
        if not linha or any(x in linha for x in ["Matrícula", "Seq.", "Origem do Vínculo"]):
            break
        
        # Regex 5 campos: Competência + Data + Contribuição + Salário + Indicadores
        # Captura: Competência, Contribuição (ignora Data e Salário), Indicadores
        # Até 2 competências por linha
        padroes = re.findall(
            r'(\\d{2}/\\d{4})\\s+\\d{2}/\\d{2}/\\d{4}\\s+([\\d.,]+)\\s+([\\d.,]+)\\s*([^\\d/]*?)(?=\\d{2}/\\d{4}|$)',
            linha
        )
        
        if padroes:
            for competencia, contribuicao, _salario_contrib, indicadores in padroes:
                contribuicao_limpa = limpar_remuneracao(contribuicao)
                
                try:
                    valor = float(contribuicao_limpa)
                    if validar_valor(valor):
                        registros.append({
                            "pagina": pagina_idx,
                            "seq": seq,
                            "codigo_emp": "FACULTATIVO",  # Sem CNPJ
                            "competencia": competencia,
                            "remuneracao": contribuicao_limpa,  # Contribuição como remuneração
                            "indicadores": indicadores.strip() or "PREC-FACULTCONC",
                        })
                except ValueError:
                    # Valor inválido, ignorar
                    pass
        
        i += 1
