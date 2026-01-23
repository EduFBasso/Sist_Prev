"""Debug do processamento CLT"""

import pdfplumber
import re
from extrator.tipos.coordenador_remuneracoes import _extrair_blocos_vinculos
from extrator.tipos.clt import processar_remuneracoes_clt

PDF_TESTE = "cnis/CNIS_JOAO_CARLOS.pdf"

with pdfplumber.open(PDF_TESTE) as pdf:
    # Pegar primeiro bloco da primeira página
    texto = pdf.pages[0].extract_text()
    blocos = _extrair_blocos_vinculos(texto)
    
    if blocos:
        bloco = blocos[0]
        print(f"📋 Debug do Bloco 1 (Seq {bloco['seq']})")
        print()
        
        print(f"CNPJ: {bloco['cnpj']}")
        print()
        
        print("─" * 70)
        print("TEXTO COMPLETO DO BLOCO:")
        print("─" * 70)
        print(bloco['texto'])
        print()
        
        print("─" * 70)
        print("PROCURANDO SEÇÃO REMUNERAÇÕES:")
        print("─" * 70)
        
        match_remun = re.search(r'Remunerações\s*(.*?)(?:Contribuições|Matrícula|$)', 
                               bloco['texto'], re.DOTALL | re.IGNORECASE)
        
        if match_remun:
            secao = match_remun.group(1)
            print(f"✅ Seção encontrada ({len(secao)} chars)")
            print()
            print("SEÇÃO REMUNERAÇÕES:")
            print(secao)
            print()
            
            # Processar
            registros = []
            processar_remuneracoes_clt(secao, bloco['seq'], bloco['cnpj'], 1, registros)
            
            print(f"Registros extraídos: {len(registros)}")
            for reg in registros[:3]:
                print(f"  {reg}")
        else:
            print("❌ Seção 'Remunerações' NÃO encontrada!")
