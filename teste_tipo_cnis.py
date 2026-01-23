"""
Teste do Detector de Tipo de CNIS

Valida se o detector identifica corretamente:
    1. Tipo de documento: EXTRATO_PREVIDENCIARIO
    2. Gatilhos presentes: INSS, CNIS, Extrato Previdenciário
"""

import pdfplumber
from extrator.tipos.tipo_cnis import (
    detectar_tipo_cnis, 
    TipoCNIS, 
    obter_nome_tipo_cnis,
    validar_tipo_cnis
)

PDF_TESTE = "cnis/CNIS_JOAO_CARLOS.pdf"

def main():
    print("=" * 70)
    print("🔍 TESTE: DETECTOR DE TIPO DE CNIS")
    print("=" * 70)
    print()
    
    print(f"📄 Arquivo: {PDF_TESTE}")
    print()
    
    with pdfplumber.open(PDF_TESTE) as pdf:
        print("─" * 70)
        print("1️⃣  DETECÇÃO AUTOMÁTICA")
        print("─" * 70)
        
        tipo_detectado = detectar_tipo_cnis(pdf)
        nome_tipo = obter_nome_tipo_cnis(tipo_detectado)
        
        print(f"Tipo detectado: {tipo_detectado.value}")
        print(f"Nome legível: {nome_tipo}")
        print()
        
        if tipo_detectado == TipoCNIS.EXTRATO_PREVIDENCIARIO:
            print("✅ Tipo identificado: EXTRATO PREVIDENCIÁRIO")
        elif tipo_detectado == TipoCNIS.DESCONHECIDO:
            print("❌ Tipo não reconhecido!")
        print()
        
        print("─" * 70)
        print("2️⃣  VALIDAÇÃO DE TIPO ESPERADO")
        print("─" * 70)
        
        # Validar tipo esperado
        tipo_esperado = TipoCNIS.EXTRATO_PREVIDENCIARIO
        valido, mensagem = validar_tipo_cnis(pdf, tipo_esperado)
        
        print(mensagem)
        print()
        
        print("─" * 70)
        print("3️⃣  ANÁLISE DE CABEÇALHO (GATILHOS)")
        print("─" * 70)
        
        # Extrair primeiras linhas para mostrar gatilhos
        primeira_pagina = pdf.pages[0].extract_text()
        linhas = primeira_pagina.split('\n')[:10]
        
        print("Primeiras 10 linhas do PDF:")
        for i, linha in enumerate(linhas, 1):
            # Destacar gatilhos
            linha_dest = linha
            if 'INSS' in linha.upper():
                linha_dest = f"🎯 {linha} [GATILHO: INSS]"
            elif 'CNIS' in linha.upper():
                linha_dest = f"🎯 {linha} [GATILHO: CNIS]"
            elif 'EXTRATO' in linha.upper() and 'PREVIDENCIÁRIO' in linha.upper():
                linha_dest = f"🎯 {linha} [GATILHO: EXTRATO PREVIDENCIÁRIO]"
            
            print(f"  {i:2d}. {linha_dest}")
        print()
        
        print("─" * 70)
        print("4️⃣  RESUMO DA VALIDAÇÃO")
        print("─" * 70)
        
        if valido and tipo_detectado == TipoCNIS.EXTRATO_PREVIDENCIARIO:
            print("✅ ESTRUTURA OK PARA PROCESSAR!")
            print()
            print("Responsabilidade 1: ✅ Identificar Tipo de CNIS")
            print("  • Tipo: Extrato Previdenciário")
            print("  • Gatilhos: INSS → CNIS → Extrato Previdenciário")
            print("  • Expansível: Sim (adicionar novos tipos em TipoCNIS enum)")
            print()
            print("Próximo passo: Validar detector de Tipos de Vínculos (Responsabilidade 2)")
        else:
            print("❌ VALIDAÇÃO FALHOU!")
            print(f"Motivo: {mensagem}")
        
    print()
    print("=" * 70)

if __name__ == "__main__":
    main()
