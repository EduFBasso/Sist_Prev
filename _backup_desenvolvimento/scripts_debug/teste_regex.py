import re

# Simulando as linhas que você descreveu
linhas_teste = [
    "01/1998 1.142,00    02/1998  1.100,00   03/1998 1.100,00",
    "04/1998 1.100,00    05/1998 1.156,00   06/1998 1.103,74",
    "07/1998 1.103,74    08/1998 1.087,86   09/1998 1.242,49   10/1998 1.432,74   11/1998 572,37",
]

for linha in linhas_teste:
    print(f"\nLinha: {linha}")
    padroes = re.findall(r'(\d{2}/\d{4})\s+([\d.,]+)\s*([^\d/]*?)(?=\d{2}/\d{4}|$)', linha)
    print(f"Encontrados: {len(padroes)}")
    for comp, valor, indic in padroes:
        print(f"  {comp} -> {valor} ({indic if indic else 'vazio'})")
