import pdfplumber

pdf_path = "cnis/CNIS_JOAO_CARLOS.pdf"

with pdfplumber.open(pdf_path) as pdf:
    pag7 = pdf.pages[6]
    texto = pag7.extract_text()
    
    print("📄 PÁGINA 7 - PROCURANDO SEQ 13:")
    print("="*70)
    
    linhas = texto.split('\n')
    for i, linha in enumerate(linhas):
        if '13 125' in linha or 'Seq. NIT Origem' in linha:
            print(f"\nLinha {i}: {linha}")
            # Mostrar contexto (10 linhas após)
            for j in range(i+1, min(i+11, len(linhas))):
                print(f"  {j}: {linhas[j]}")
            break
    
    print("\n🔍 PROCURANDO '11/2025' NA PÁGINA 7:")
    print("="*70)
    for i, linha in enumerate(linhas):
        if '11/2025' in linha:
            print(f"Linha {i}: {linha}")
            if i > 0:
                print(f"  Anterior: {linhas[i-1]}")
            if i+1 < len(linhas):
                print(f"  Próxima: {linhas[i+1]}")
