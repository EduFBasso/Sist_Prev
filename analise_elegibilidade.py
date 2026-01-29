"""
Módulo de Análise de Elegibilidade Previdenciária

Funções para calcular e analisar aposentadorias:
- Regra Pré-Reforma (até 13/11/2019)
- Regra Pós-Reforma (EC 103/2019)
- Simulações e comparações
"""

from datetime import datetime, date
from dateutil.relativedelta import relativedelta
from typing import Dict, Tuple, Optional


# ============================================================================
# CONSTANTES
# ============================================================================

DATA_REFORMA = date(2019, 11, 13)
SALARIO_MINIMO_2024 = 1412.00  # Atualizar conforme necessário
TETO_INSS_2024 = 7786.02  # Atualizar conforme necessário


# ============================================================================
# FUNÇÕES DE CÁLCULO DE TEMPO
# ============================================================================

def calcular_idade(data_nascimento: str, data_referencia: Optional[date] = None) -> Dict:
    """
    Calcula idade em anos, meses e dias.
    
    Args:
        data_nascimento: String no formato DD/MM/AAAA
        data_referencia: Data de referência (default: hoje)
    
    Returns:
        Dict com anos, meses, dias, total_meses
    """
    if data_referencia is None:
        data_referencia = date.today()
    
    # Converter string para date
    try:
        if isinstance(data_nascimento, str):
            partes = data_nascimento.split('/')
            nasc = date(int(partes[2]), int(partes[1]), int(partes[0]))
        else:
            nasc = data_nascimento
    except:
        return {'anos': 0, 'meses': 0, 'dias': 0, 'total_meses': 0}
    
    diff = relativedelta(data_referencia, nasc)
    
    return {
        'anos': diff.years,
        'meses': diff.months,
        'dias': diff.days,
        'total_meses': diff.years * 12 + diff.months
    }


def calcular_tempo_contribuicao(total_competencias: int) -> Dict:
    """
    Calcula tempo de contribuição a partir do número de competências.
    
    Args:
        total_competencias: Número de meses contribuídos
    
    Returns:
        Dict com anos, meses, total_meses
    """
    anos = total_competencias // 12
    meses = total_competencias % 12
    
    return {
        'anos': anos,
        'meses': meses,
        'total_meses': total_competencias
    }


# ============================================================================
# FUNÇÕES DE CÁLCULO - PRÉ-REFORMA
# ============================================================================

def calcular_fator_previdenciario(tempo_contrib_anos: float, idade_anos: float, sexo: str) -> float:
    """
    Calcula fator previdenciário (regra antiga - Lei 9.876/99).
    
    Fórmula: f = (Tc × a / Es) × [1 + (Id + Tc × a) / 100]
    
    Onde:
    - Tc = tempo de contribuição em anos
    - a = alíquota de contribuição (0.31)
    - Id = idade no momento da aposentadoria
    - Es = expectativa de sobrevida (tabela IBGE)
    
    Args:
        tempo_contrib_anos: Tempo de contribuição em anos (decimal)
        idade_anos: Idade em anos (decimal)
        sexo: 'Masculino' ou 'Feminino'
    
    Returns:
        Fator previdenciário (float entre ~0.5 e 1.2)
    """
    Tc = tempo_contrib_anos
    Id = idade_anos
    a = 0.31  # Alíquota de contribuição
    
    # Tabela IBGE de expectativa de sobrevida (simplificada)
    # Valores aproximados para diferentes idades
    idade_int = int(Id)
    
    if sexo == 'Masculino':
        # Expectativa homens (IBGE 2023)
        tabela_es = {
            40: 38.5, 45: 33.9, 50: 29.5, 51: 28.6, 52: 27.7, 53: 26.8,
            54: 26.0, 55: 25.1, 56: 24.3, 57: 23.4, 58: 22.6, 59: 21.8,
            60: 21.0, 61: 20.2, 62: 19.4, 63: 18.7, 64: 17.9, 65: 17.2
        }
    else:
        # Expectativa mulheres (IBGE 2023)
        tabela_es = {
            40: 43.2, 45: 38.4, 50: 33.8, 51: 32.9, 52: 32.0, 53: 31.1,
            54: 30.2, 55: 29.3, 56: 28.5, 57: 27.6, 58: 26.7, 59: 25.9,
            60: 25.0, 61: 24.2, 62: 23.4, 63: 22.6, 64: 21.8, 65: 21.0
        }
    
    # Buscar expectativa de sobrevida (interpolar se necessário)
    if idade_int in tabela_es:
        Es = tabela_es[idade_int]
    elif idade_int < 40:
        Es = 40.0  # Aproximação para idades menores
    elif idade_int > 65:
        Es = 15.0  # Aproximação para idades maiores
    else:
        # Interpolar linearmente
        idade_menor = (idade_int // 5) * 5
        idade_maior = idade_menor + 5
        es_menor = tabela_es.get(idade_menor, 20.0)
        es_maior = tabela_es.get(idade_maior, 20.0)
        Es = es_menor + (es_maior - es_menor) * ((idade_int - idade_menor) / 5)
    
    # Fórmula do fator previdenciário (Lei 9.876/99)
    # f = (Tc × a / Es) × [1 + (Id + Tc × a) / 100]
    parte1 = (Tc * a) / Es
    parte2 = 1 + ((Id + (Tc * a)) / 100)
    
    fator = parte1 * parte2
    
    # Garantir que fator está em faixa razoável (0.4 a 1.3)
    fator = max(0.4, min(fator, 1.3))
    
    return round(fator, 4)


def calcular_media_80_maiores(valores_corrigidos: list) -> float:
    """
    Calcula média dos 80% maiores salários (regra antiga).
    
    Args:
        valores_corrigidos: Lista de valores corrigidos monetariamente
    
    Returns:
        Média dos 80% maiores
    """
    if not valores_corrigidos:
        return 0.0
    
    # Ordenar decrescente
    valores_ordenados = sorted(valores_corrigidos, reverse=True)
    
    # Pegar 80%
    qtd_80_porcento = int(len(valores_ordenados) * 0.8)
    if qtd_80_porcento == 0:
        qtd_80_porcento = len(valores_ordenados)
    
    valores_80 = valores_ordenados[:qtd_80_porcento]
    
    # Calcular média
    media = sum(valores_80) / len(valores_80)
    
    return round(media, 2)


def verificar_elegibilidade_pre_reforma(
    tempo_contrib_meses: int,
    idade_anos: float,
    sexo: str,
    data_referencia: Optional[date] = None
) -> Dict:
    """
    Verifica elegibilidade pela regra antiga (pré-reforma).
    
    Args:
        tempo_contrib_meses: Tempo de contribuição total em meses
        idade_anos: Idade atual em anos
        sexo: 'Masculino' ou 'Feminino'
        data_referencia: Data de verificação (default: hoje)
    
    Returns:
        Dict com análise completa
    """
    if data_referencia is None:
        data_referencia = date.today()
    
    tempo_anos = tempo_contrib_meses / 12
    
    # Requisitos pré-reforma
    tempo_necessario_meses = 420 if sexo == 'Masculino' else 360  # 35H / 30M
    tempo_necessario_anos = tempo_necessario_meses / 12
    
    # Verificar direito adquirido (antes da reforma)
    tem_direito_adquirido = data_referencia < DATA_REFORMA and tempo_contrib_meses >= tempo_necessario_meses
    
    # Elegibilidade atual
    elegivel = tempo_contrib_meses >= tempo_necessario_meses
    
    # Tempo faltante
    falta_meses = max(0, tempo_necessario_meses - tempo_contrib_meses)
    
    # Calcular fator previdenciário
    if elegivel:
        fator = calcular_fator_previdenciario(tempo_anos, idade_anos, sexo)
    else:
        # Simular fator quando atingir requisitos
        tempo_futuro = tempo_necessario_anos
        idade_futura = idade_anos + (falta_meses / 12)
        fator = calcular_fator_previdenciario(tempo_futuro, idade_futura, sexo)
    
    return {
        'elegivel': elegivel,
        'direito_adquirido': tem_direito_adquirido,
        'tempo_atual_meses': tempo_contrib_meses,
        'tempo_atual_anos': round(tempo_anos, 2),
        'tempo_necessario_meses': tempo_necessario_meses,
        'tempo_necessario_anos': tempo_necessario_anos,
        'falta_meses': falta_meses,
        'falta_anos': round(falta_meses / 12, 2),
        'idade_atual': idade_anos,
        'fator_previdenciario': fator,
        'observacao': 'Regra extinta pela EC 103/2019 - válida apenas para direito adquirido'
    }


# ============================================================================
# FUNÇÕES DE CÁLCULO - PÓS-REFORMA
# ============================================================================

def calcular_coeficiente_pos_reforma(tempo_contrib_meses: int, sexo: str) -> Dict:
    """
    Calcula coeficiente da regra nova (pós-reforma).
    
    Base: 60%
    Acréscimo: 2% por ano acima de 20H / 15M
    Máximo: 100%
    
    Args:
        tempo_contrib_meses: Tempo de contribuição em meses
        sexo: 'Masculino' ou 'Feminino'
    
    Returns:
        Dict com coeficiente e detalhes
    """
    tempo_anos = tempo_contrib_meses / 12
    
    # Tempo base para acréscimo
    tempo_base_anos = 20 if sexo == 'Masculino' else 15
    
    # Coeficiente base
    coef_base = 0.60  # 60%
    
    # Anos acima do base
    anos_acima = max(0, tempo_anos - tempo_base_anos)
    
    # Acréscimo (2% por ano)
    acrescimo = anos_acima * 0.02
    
    # Coeficiente total
    coef_total = coef_base + acrescimo
    
    # Limitar a 100%
    coef_final = min(coef_total, 1.0)
    
    return {
        'coeficiente_base': coef_base,
        'anos_acima_base': round(anos_acima, 2),
        'acrescimo': round(acrescimo, 4),
        'coeficiente_total': round(coef_total, 4),
        'coeficiente_final': round(coef_final, 4),
        'percentual': round(coef_final * 100, 2)
    }


def verificar_elegibilidade_pos_reforma(
    tempo_contrib_meses: int,
    idade_anos: float,
    sexo: str
) -> Dict:
    """
    Verifica elegibilidade pela regra nova (pós-reforma EC 103/2019).
    
    Requisitos:
    - Idade mínima: 65H / 62M
    - Tempo mínimo: 15 anos (180 meses) - CARÊNCIA
    - Para coeficiente 60%: 20H / 15M
    - Para 100%: 35H / 30M
    
    Args:
        tempo_contrib_meses: Tempo de contribuição total em meses
        idade_anos: Idade atual em anos
        sexo: 'Masculino' ou 'Feminino'
    
    Returns:
        Dict com análise completa
    """
    # Requisitos mínimos
    idade_minima = 65 if sexo == 'Masculino' else 62
    tempo_minimo_meses = 180  # 15 anos - CARÊNCIA MÍNIMA (EC 103/2019)
    tempo_60_meses = 240 if sexo == 'Masculino' else 180  # 20H / 15M - para coef 60%
    tempo_100_meses = 420 if sexo == 'Masculino' else 360  # 35H / 30M
    
    # Verificar requisitos
    atende_idade = idade_anos >= idade_minima
    atende_tempo = tempo_contrib_meses >= tempo_minimo_meses
    
    elegivel = atende_idade and atende_tempo
    
    # Tempo faltante
    falta_idade_anos = max(0, idade_minima - idade_anos)
    falta_tempo_meses = max(0, tempo_minimo_meses - tempo_contrib_meses)
    falta_60_meses = max(0, tempo_60_meses - tempo_contrib_meses)
    falta_100_meses = max(0, tempo_100_meses - tempo_contrib_meses)
    
    # Calcular coeficiente
    coef_info = calcular_coeficiente_pos_reforma(tempo_contrib_meses, sexo)
    
    return {
        'elegivel': elegivel,
        'atende_idade': atende_idade,
        'atende_tempo': atende_tempo,
        'idade_atual': idade_anos,
        'idade_minima': idade_minima,
        'falta_idade_anos': round(falta_idade_anos, 2),
        'tempo_atual_meses': tempo_contrib_meses,
        'tempo_minimo_meses': tempo_minimo_meses,
        'falta_tempo_meses': falta_tempo_meses,
        'tempo_60_meses': tempo_60_meses,
        'falta_60_meses': falta_60_meses,
        'tempo_100_meses': tempo_100_meses,
        'falta_100_meses': falta_100_meses,
        'coeficiente': coef_info['coeficiente_final'],
        'coeficiente_percentual': coef_info['percentual'],
        'observacao': 'Regra atual (EC 103/2019) - Carência 15 anos'
    }


# ============================================================================
# SIMULAÇÕES
# ============================================================================

def simular_contribuicao_facultativa(
    tempo_atual_meses: int,
    media_atual: float,
    valor_mensal: float,
    meses_contribuir: int
) -> Dict:
    """
    Simula impacto de contribuições facultativas futuras.
    
    Args:
        tempo_atual_meses: Tempo atual de contribuição
        media_atual: Média salarial atual
        valor_mensal: Valor da contribuição mensal
        meses_contribuir: Quantos meses vai contribuir
    
    Returns:
        Dict com simulação
    """
    tempo_final_meses = tempo_atual_meses + meses_contribuir
    
    # Nova média (simplificado - assume INPC constante)
    soma_atual = media_atual * tempo_atual_meses
    soma_nova = soma_atual + (valor_mensal * meses_contribuir)
    nova_media = soma_nova / tempo_final_meses
    
    # Data prevista
    data_hoje = date.today()
    data_prevista = data_hoje + relativedelta(months=meses_contribuir)
    
    return {
        'tempo_final_meses': tempo_final_meses,
        'tempo_final_anos': round(tempo_final_meses / 12, 2),
        'nova_media': round(nova_media, 2),
        'custo_total': valor_mensal * meses_contribuir,
        'data_prevista': data_prevista.strftime('%d/%m/%Y'),
        'valor_mensal': valor_mensal
    }


def calcular_valor_para_manter_media(tempo_atual_meses: int, media_atual: float) -> float:
    """
    Calcula valor de contribuição necessário para manter a média atual.
    
    Args:
        tempo_atual_meses: Tempo atual em meses
        media_atual: Média salarial atual
    
    Returns:
        Valor mensal necessário
    """
    # Para manter média: novo_valor = média_atual
    return round(media_atual, 2)
