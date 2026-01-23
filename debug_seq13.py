import pdfplumber
from extrator.tipos.coordenador_remuneracoes import _extrair_zona_util, _extrair_blocos_vinculos
from extrator.tipos.detector import detectar_tipo_vinculo

pdf_path = "cnis/CNIS_JOAO_CARLOS.pdf"

with pdfplumber.open(pdf_path) as pdf:
    pag7 = pdf.pages[6]
    texto7 = pag7.extract_text()
    
    print("🔍 PROCURANDO SEQ 13 NA PÁGINA 7")
    print("="*70)
    
    # Ver se tem "13 125"
    if "13 125" in texto7:
        print("✅ Seq 13 ENCONTRADA na página 7\n")
        
        # Extrair zona útil
        zona_util = _extrair_zona_util(texto7)
        
        # Extrair blocos
        blocos = _extrair_blocos_vinculos(zona_util)
        
        print(f"📦 BLOCOS EXTRAÍDOS: {len(blocos)}")
        
        for i, bloco in enumerate(blocos, 1):
            if bloco['seq'] == '13':
                print(f"\n--- BLOCO {i}: SEQ 13 ---")
                print(f"Seq: {bloco['seq']}")
                print(f"NIT: {bloco['nit']}")
                print(f"Tipo: {detectar_tipo_vinculo(bloco['texto'])}")
                
                # Ver texto do bloco
                linhas = bloco['texto'].split('\n')
                print(f"\nTexto completo ({len(linhas)} linhas):")
                for j, linha in enumerate(linhas):
                    print(f"  {j:2d}: {linha}")
                
                # Ver se tem "11/2025"
                if "11/2025" in bloco['texto']:
                    print("\n✅ TEM '11/2025' no bloco")
                else:
                    print("\n⚠️  NÃO TEM '11/2025' no bloco")
    else:
        print("❌ Seq 13 NÃO encontrada na página 7")
