import pdfplumber

with pdfplumber.open("CNIS_JOAO_CARLOS.pdf") as pdf:
    # Página 2 (índice 1)
    pagina = pdf.pages[1]
    texto = pagina.extract_text() or ""
    
    print("="*80)
    print("PÁGINA 2 - PRIMEIRAS 50 LINHAS")
    print("="*80)
    
    linhas = texto.split('\n')
    for i, linha in enumerate(linhas[:50]):
        print(f"{i:3d}: {linha}")
