import pdfplumber
from extrator.tipos.coordenador_remuneracoes import _extrair_zona_util
import re

pdf_path = "cnis/CNIS_JOAO_CARLOS.pdf"

with pdfplumber.open(pdf_path) as pdf:
    pag7 = pdf.pages[6]
    texto7 = pag7.extract_text()
    zona_util7 = _extrair_zona_util(texto7)
    
    print("🔍 PROCURANDO SEÇÃO 'Contribuições' COM CABEÇALHO:")
    print("="*70)
    
    # Tentar match 1
    match1 = re.search(r'Contribuições\s+Competência\s+Data\s+Pgto', 
                      zona_util7, re.IGNORECASE)
    print(f"Match 'Contribuições Competência Data Pgto': {match1 is not None}")
    if match1:
        print(f"  Posição: {match1.start()}")
        print(f"  Texto: {zona_util7[match1.start():match1.start()+100]}")
    
    # Tentar match 2 (fallback)
    match2 = re.search(r'(?:^|\n)Competência\s+Data\s+Pgto',
                      zona_util7[:300], re.IGNORECASE)
    print(f"\nMatch fallback 'Competência Data Pgto' (300 chars): {match2 is not None}")
    if match2:
        print(f"  Posição: {match2.start()}")
        print(f"  Texto: {zona_util7[match2.start():match2.start()+100]}")
    
    print(f"\n📝 PRIMEIROS 400 CHARS DA ZONA ÚTIL:")
    print("="*70)
    print(zona_util7[:400])
