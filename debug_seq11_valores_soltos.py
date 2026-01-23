import pdfplumber
from extrator.tipos.coordenador_remuneracoes import _extrair_zona_util

pdf_path = "cnis/CNIS_JOAO_CARLOS.pdf"

with pdfplumber.open(pdf_path) as pdf:
    pag6 = pdf.pages[5]
    pag7 = pdf.pages[6]
    
    texto6 = pag6.extract_text()
    texto7 = pag7.extract_text()
    
    print("📄 PÁGINA 6 - ÚLTIMAS 5 LINHAS:")
    print("="*70)
    linhas6 = texto6.split('\n')
    for linha in linhas6[-5:]:
        print(linha)
    
    print("\n📄 PÁGINA 7 - ZONA ÚTIL (PRIMEIRAS 300 CHARS):")
    print("="*70)
    zona_util7 = _extrair_zona_util(texto7)
    print(zona_util7[:300])
    
    print("\n🔍 ANÁLISE:")
    print("="*70)
    print("Seq 11 na página 6: RECOLHIMENTO Facultativo")
    print("Valores órfãos na página 7: 09/2019 e 10/2019")
    print("ANTES do primeiro 'Seq. NIT Origem do Vínculo'")
    print("\n➡️  São 'valores soltos' de contribuições facultativas!")
