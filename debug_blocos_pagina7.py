import pdfplumber
from extrator.tipos.coordenador_remuneracoes import _extrair_zona_util, _extrair_blocos_vinculos

pdf_path = "cnis/CNIS_JOAO_CARLOS.pdf"

with pdfplumber.open(pdf_path) as pdf:
    pagina7 = pdf.pages[6]  # página 7 (índice 6)
    texto = pagina7.extract_text()
    
    # Extrair zona útil
    zona_util = _extrair_zona_util(texto)
    
    print(f"📄 ZONA ÚTIL PÁGINA 7:")
    print("="*80)
    print(zona_util[:500])
    print("...")
    print("="*80)
    
    # Extrair blocos
    blocos = _extrair_blocos_vinculos(zona_util)
    
    print(f"\n📦 BLOCOS EXTRAÍDOS: {len(blocos)}")
    for i, bloco in enumerate(blocos):
        print(f"\n--- BLOCO {i+1} ---")
        print(f"Primeiras 10 linhas:")
        linhas = bloco.split('\n')[:10]
        for linha in linhas:
            print(f"  {linha}")
        print(f"...")
