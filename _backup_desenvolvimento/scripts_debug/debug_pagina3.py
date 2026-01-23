import pdfplumber
from pathlib import Path

pdf_path = Path("CNIS_JOAO_CARLOS.pdf")

with pdfplumber.open(pdf_path) as pdf:
    # Página 3 é onde está a continuação da Seq 2
    pagina = pdf.pages[2]  # índice 2 = página 3
    texto = pagina.extract_text() or ""
    
    print("="*80)
    print("PÁGINA 3 - TEXTO COMPLETO")
    print("="*80)
    print(texto[:2000])  # Primeiros 2000 caracteres
    
    print("\n" + "="*80)
    print("LINHAS COM 1998")
    print("="*80)
    
    for i, linha in enumerate(texto.split('\n')):
        if '1998' in linha:
            print(f"Linha {i}: {linha}")
