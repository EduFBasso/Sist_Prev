import pdfplumber
from extrator.tipos.coordenador_remuneracoes import _extrair_zona_util, _extrair_blocos_vinculos
from extrator.tipos.detector import detectar_tipo_vinculo

pdf_path = "cnis/CNIS_JOAO_CARLOS.pdf"

with pdfplumber.open(pdf_path) as pdf:
    pagina7 = pdf.pages[6]  # página 7 (índice 6)
    texto = pagina7.extract_text()
    
    zona_util = _extrair_zona_util(texto)
    blocos = _extrair_blocos_vinculos(zona_util)
    
    print(f"📦 BLOCOS EXTRAÍDOS NA PÁGINA 7: {len(blocos)}")
    print()
    
    for i, bloco in enumerate(blocos, 1):
        print(f"--- BLOCO {i} ---")
        print(f"Seq: {bloco.get('seq', '?')}")
        print(f"NIT: {bloco.get('nit', '?')}")
        print(f"CNPJ: {bloco.get('cnpj', '?')}")
        
        # Detectar tipo
        tipo = detectar_tipo_vinculo(bloco['texto'])
        print(f"Tipo detectado: {tipo}")
        
        # Ver primeiras linhas
        linhas = bloco['texto'].split('\n')[:8]
        print(f"\nPrimeiras 8 linhas:")
        for linha in linhas:
            print(f"  {linha}")
        print()
