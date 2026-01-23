#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Ver estrutura real da página 7 - Contribuições"""

import pdfplumber
import re

pdf = pdfplumber.open('cnis/CNIS_JOAO_CARLOS.pdf')
page7 = pdf.pages[6].extract_text()

# Simular zona útil
inicio_zona = page7.find("Relações Previdenciárias")
fim_zona = page7.find("O INSS poderá rever")
zona_util = page7[inicio_zona:fim_zona]

print("=== Buscando 'Origem do Vínculo' ===")
pos = 0
count = 0
while True and count < 10:  # Limite para evitar loop infinito no teste
    idx_origem = zona_util.find("Origem do Vínculo", pos)
    if idx_origem == -1:
        break
    
    print(f"\nOcorrência {count+1} at position {idx_origem}")
    
    # Extrair próximas 3 linhas
    fim_linha = zona_util.find('\n', idx_origem)
    linha1_inicio = fim_linha + 1
    linha1_fim = zona_util.find('\n', linha1_inicio)
    linha1 = zona_util[linha1_inicio:linha1_fim]
    
    print(f"Próxima linha: {linha1[:100]}")
    
    # Verificar se tem NIT + RECOLHIMENTO
    tem_nit = bool(re.search(r'\d{3}\.\d{5}\.\d{2}-\d', linha1))
    tem_recol = "RECOLHIMENTO" in linha1.upper()
    print(f"  Tem NIT: {tem_nit}, Tem RECOLHIMENTO: {tem_recol}")
    
    if tem_nit and tem_recol:
        # Extrair Seq
        seq_match = re.match(r'^\s*(\d+)\s+\d{3}\.\d{5}\.\d{2}-\d', linha1)
        if seq_match:
            print(f"  ✓ Seq encontrado: {seq_match.group(1)}")
    
    # Avançar posição
    pos = linha1_fim
    count += 1

print(f"\n\nTotal de ocorrências: {count}")

pdf_path = "cnis/CNIS_JOAO_CARLOS.pdf"

with pdfplumber.open(pdf_path) as pdf:
    # Página 7
    texto = pdf.pages[6].extract_text()  # índice 6 = página 7
    linhas = texto.split('\n')
    
    print("="*70)
    print("PÁGINA 7 - Primeiras 40 linhas")
    print("="*70)
    
    for i, linha in enumerate(linhas[:40], 1):
        print(f"{i:2d}: {linha}")

