#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Debug Seq 11 - Contribuições Facultativas"""

import pdfplumber
import re

pdf_path = "cnis/CNIS_JOAO_CARLOS.pdf"

with pdfplumber.open(pdf_path) as pdf:
    # Páginas 6 e 7
    for pag_num in [6, 7]:
        pagina = pdf.pages[pag_num - 1]
        texto = pagina.extract_text() or ""
        linhas = texto.split('\n')
        
        print(f"\n{'='*70}")
        print(f"PÁGINA {pag_num}")
        print('='*70)
        
        # Procurar "Seq." com "11"
        tem_seq11 = False
        for i, linha in enumerate(linhas):
            if re.search(r'Seq\.', linha):
                print(f"\n🔍 Linha {i} com 'Seq.': {linha}")
                # Mostrar próximas 3 linhas
                for j in range(i+1, min(i+4, len(linhas))):
                    print(f"   +{j-i}: {linhas[j]}")
                if '11' in '\n'.join(linhas[i:i+3]):
                    tem_seq11 = True
        
        # Procurar "Contribuições"
        for i, linha in enumerate(linhas):
            if "Contribuições" in linha or "Contribuicoes" in linha:
                print(f"\n✅ Linha {i} com 'Contribuições': {linha}")
                # Mostrar próximas 10 linhas
                print("Próximas linhas:")
                for j in range(i+1, min(i+11, len(linhas))):
                    print(f"   {j}: {linhas[j]}")
        
        print(f"\n📊 Tem Seq 11: {'SIM' if tem_seq11 else 'NÃO'}")


