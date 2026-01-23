#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Testar regex com linha real da página 7"""

import re

# Linha 12 da página 7:
linha = "09/2019 18/09/2019 200,00 1.000,00 PREC-FACULTCONC 10/2019 22/11/2019 199,60 998,00 PREC-FACULTCONC"

print("Linha a processar:")
print(linha)
print()

# Regex atual:
regex = r'(\d{2}/\d{4})\s+\d{2}/\d{2}/\d{4}\s+([\d.,]+)\s+([\d.,]+)\s+([A-Z\-]+)'

matches = re.findall(regex, linha)

print(f"Regex: {regex}")
print(f"Matches encontrados: {len(matches)}")
print()

for i, match in enumerate(matches, 1):
    comp, contrib, salario, indic = match
    print(f"Match {i}:")
    print(f"  Competência: {comp}")
    print(f"  Contribuição: {contrib}")
    print(f"  Salário: {salario}")
    print(f"  Indicadores: {indic}")
