import pdfplumber

pdf_path = "cnis/CNIS_JOAO_CARLOS.pdf"

with pdfplumber.open(pdf_path) as pdf:
    pagina7 = pdf.pages[6]  # página 7 (índice 6)
    texto = pagina7.extract_text()
    
    print(f"📄 PÁGINA 7 - TEXTO COMPLETO:")
    print("="*80)
    print(texto)
    print("="*80)
    
    # Procurar vínculos facultativos
    linhas = texto.split('\n')
    print(f"\n🔍 PROCURANDO PADRÕES FACULTATIVOS...")
    for i, linha in enumerate(linhas):
        if 'Facultativo' in linha or 'FACULTATIVO' in linha:
            print(f"\nLinha {i}: {linha}")
            # Contexto ao redor
            if i > 0:
                print(f"  Anterior: {linhas[i-1]}")
            if i+1 < len(linhas):
                print(f"  Próxima: {linhas[i+1]}")
