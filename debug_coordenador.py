"""Debug do coordenador - ver o que está sendo extraído"""

import pdfplumber
from extrator.tipos.coordenador_remuneracoes import _extrair_blocos_vinculos

PDF_TESTE = "cnis/CNIS_JOAO_CARLOS.pdf"

with pdfplumber.open(PDF_TESTE) as pdf:
    print("📄 Testando extração de blocos...")
    print()
    
    for pag_idx, pagina in enumerate(pdf.pages[:3], 1):  # Primeiras 3 páginas
        texto = pagina.extract_text()
        blocos = _extrair_blocos_vinculos(texto)
        
        print(f"Página {pag_idx}: {len(blocos)} blocos")
        
        for i, bloco in enumerate(blocos, 1):
            print(f"  Bloco {i}:")
            print(f"    Seq: {bloco.get('seq')}")
            print(f"    NIT: {bloco.get('nit')}")
            print(f"    CNPJ: {bloco.get('cnpj')}")
            print(f"    Texto: {len(bloco.get('texto', ''))} chars")
            
            # Mostrar primeiras 5 linhas
            linhas = bloco.get('texto', '').split('\n')[:5]
            for linha in linhas:
                print(f"      {linha[:70]}")
        print()
