import pdfplumber

pdf_path = "cnis/CNIS_JOAO_CARLOS.pdf"

with pdfplumber.open(pdf_path) as pdf:
    # Página 6
    pagina6 = pdf.pages[5]
    texto6 = pagina6.extract_text()
    
    print(f"📄 PÁGINA 6 - ÚLTIMAS 30 LINHAS:")
    print("="*80)
    linhas6 = texto6.split('\n')
    for linha in linhas6[-30:]:
        print(linha)
    print("="*80)
    
    # Página 7
    pagina7 = pdf.pages[6]
    texto7 = pagina7.extract_text()
    
    print(f"\n📄 PÁGINA 7 - PRIMEIRAS 15 LINHAS:")
    print("="*80)
    linhas7 = texto7.split('\n')
    for linha in linhas7[:15]:
        print(linha)
    print("="*80)
