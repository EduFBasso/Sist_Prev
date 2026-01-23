import pdfplumber

with pdfplumber.open("CNIS_JOAO_CARLOS.pdf") as pdf:
    for i, pagina in enumerate(pdf.pages, start=1):
        texto = pagina.extract_text() or ""
        
        # Procurar por 12/1997 (última competência conhecida da Seq 2)
        if '12/1997' in texto:
            print(f"\n{'='*80}")
            print(f"PÁGINA {i} - Contém 12/1997")
            print('='*80)
            
            linhas = texto.split('\n')
            for j, linha in enumerate(linhas):
                if '12/1997' in linha or '01/1998' in linha:
                    print(f"Linha {j}: {linha}")
                    # Mostrar 5 linhas seguintes
                    for k in range(1, 6):
                        if j + k < len(linhas):
                            print(f"  +{k}: {linhas[j+k]}")
