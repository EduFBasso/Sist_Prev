"""
Detector de Tipo de Documento CNIS

RESPONSABILIDADE: Identificar o tipo de documento CNIS pela estrutura do cabeçalho.

TIPOS CONHECIDOS:
    1. EXTRATO_PREVIDENCIARIO - Documento padrão atual
       Gatilhos: "INSS" → "CNIS - Cadastro Nacional de Informações Sociais" → "Extrato Previdenciário"
    
    Futuros tipos podem ser adicionados aqui conforme necessário.

FILOSOFIA:
    - Estrutura expansível: novos tipos = novo enum + novos gatilhos
    - Gatilhos claros: "tem X, não tem Y, tipo de cabeçalho Z"
    - Falha segura: retorna DESCONHECIDO se não identificar
    
EXEMPLO DE USO:
    ```python
    import pdfplumber
    from extrator.tipos.tipo_cnis import detectar_tipo_cnis, TipoCNIS
    
    with pdfplumber.open('extrato.pdf') as pdf:
        tipo = detectar_tipo_cnis(pdf)
        
        if tipo == TipoCNIS.EXTRATO_PREVIDENCIARIO:
            print("Processando extrato previdenciário padrão")
        elif tipo == TipoCNIS.DESCONHECIDO:
            print("Tipo de documento não reconhecido")
    ```
"""

from enum import Enum
import re


class TipoCNIS(Enum):
    """
    Tipos de documentos CNIS conhecidos.
    
    Adicione novos tipos aqui conforme necessário:
        NOVO_TIPO = "novo_tipo"
    """
    EXTRATO_PREVIDENCIARIO = "extrato_previdenciario"
    DESCONHECIDO = "desconhecido"


# ========================================
# GATILHOS IDENTIFICADORES
# ========================================

GATILHOS_EXTRATO_PREVIDENCIARIO = {
    'titulo_inss': re.compile(r'\bINSS\b', re.IGNORECASE),
    'titulo_cnis': re.compile(r'CNIS\s*-?\s*Cadastro Nacional de Informa[çc][õo]es Sociais', re.IGNORECASE),
    'subtitulo_extrato': re.compile(r'Extrato Previdenci[áa]rio', re.IGNORECASE),
}


def detectar_tipo_cnis(pdf) -> TipoCNIS:
    """
    Detecta o tipo de documento CNIS analisando o cabeçalho das primeiras páginas.
    
    ALGORITMO:
        1. Extrai texto das 2 primeiras páginas
        2. Verifica presença de gatilhos identificadores
        3. Retorna tipo correspondente ou DESCONHECIDO
    
    Args:
        pdf: Objeto pdfplumber.PDF
        
    Returns:
        TipoCNIS: Tipo identificado
        
    EXPANSÃO FUTURA:
        Para adicionar novo tipo:
        1. Adicionar enum em TipoCNIS
        2. Criar dicionário GATILHOS_NOVO_TIPO
        3. Adicionar verificação nesta função
    """
    # Extrair texto das 2 primeiras páginas (cabeçalho geralmente está aqui)
    texto_cabecalho = ""
    for i, pagina in enumerate(pdf.pages[:2]):
        texto_cabecalho += pagina.extract_text() or ""
        if i == 0:  # Primeira página geralmente é suficiente
            texto_cabecalho += "\n"
    
    # Verificar se é Extrato Previdenciário (tipo padrão atual)
    if _verificar_extrato_previdenciario(texto_cabecalho):
        return TipoCNIS.EXTRATO_PREVIDENCIARIO
    
    # Adicione aqui verificações para novos tipos no futuro:
    # if _verificar_novo_tipo(texto_cabecalho):
    #     return TipoCNIS.NOVO_TIPO
    
    # Tipo não identificado
    return TipoCNIS.DESCONHECIDO


def _verificar_extrato_previdenciario(texto: str) -> bool:
    """
    Verifica se o documento é um Extrato Previdenciário padrão.
    
    GATILHOS OBRIGATÓRIOS:
        ✅ "INSS" (topo esquerdo)
        ✅ "CNIS - Cadastro Nacional de Informações Sociais"
        ✅ "Extrato Previdenciário"
    
    LÓGICA:
        Todos os 3 gatilhos devem estar presentes para confirmar tipo.
    
    Args:
        texto: Texto extraído do cabeçalho do PDF
        
    Returns:
        bool: True se todos os gatilhos estão presentes
    """
    gatilhos = GATILHOS_EXTRATO_PREVIDENCIARIO
    
    tem_inss = bool(gatilhos['titulo_inss'].search(texto))
    tem_cnis = bool(gatilhos['titulo_cnis'].search(texto))
    tem_extrato = bool(gatilhos['subtitulo_extrato'].search(texto))
    
    # TODOS os gatilhos devem estar presentes
    return tem_inss and tem_cnis and tem_extrato


def obter_nome_tipo_cnis(tipo: TipoCNIS) -> str:
    """
    Retorna nome legível do tipo de CNIS.
    
    Args:
        tipo: Enum TipoCNIS
        
    Returns:
        str: Nome legível para exibição
    """
    nomes = {
        TipoCNIS.EXTRATO_PREVIDENCIARIO: "Extrato Previdenciário (CNIS Padrão)",
        TipoCNIS.DESCONHECIDO: "Tipo Desconhecido",
    }
    return nomes.get(tipo, "Tipo Não Mapeado")


# ========================================
# FUNÇÃO DE VALIDAÇÃO (ÚTIL PARA TESTES)
# ========================================

def validar_tipo_cnis(pdf, tipo_esperado: TipoCNIS) -> tuple[bool, str]:
    """
    Valida se o PDF corresponde ao tipo esperado.
    
    ÚTIL PARA:
        - Validação em pipeline de processamento
        - Testes automatizados
        - Rejeição de documentos incorretos
    
    Args:
        pdf: Objeto pdfplumber.PDF
        tipo_esperado: Tipo esperado do documento
        
    Returns:
        tuple[bool, str]: (é_válido, mensagem)
        
    Exemplo:
        ```python
        valido, msg = validar_tipo_cnis(pdf, TipoCNIS.EXTRATO_PREVIDENCIARIO)
        if not valido:
            raise ValueError(f"Documento inválido: {msg}")
        ```
    """
    tipo_detectado = detectar_tipo_cnis(pdf)
    
    if tipo_detectado == tipo_esperado:
        return True, f"✅ Documento confirmado: {obter_nome_tipo_cnis(tipo_detectado)}"
    
    if tipo_detectado == TipoCNIS.DESCONHECIDO:
        return False, f"❌ Tipo de documento não reconhecido. Esperado: {obter_nome_tipo_cnis(tipo_esperado)}"
    
    return False, f"❌ Tipo incorreto. Detectado: {obter_nome_tipo_cnis(tipo_detectado)}, Esperado: {obter_nome_tipo_cnis(tipo_esperado)}"
