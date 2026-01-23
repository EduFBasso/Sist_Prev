"""Debug da Seq 10 - verificar onde está"""

import pdfplumber
from extrator.tipos.coordenador_remuneracoes import (
    _extrair_blocos_vinculos,
    _extrair_zona_util
)

PDF_TESTE = "cnis/CNIS_JOAO_CARLOS.pdf"

with pdfplumber.open(PDF_TESTE) as pdf:
    print("🔍 PROCURANDO SEQ 10")
    print("=" * 70)
    print()
    
    for pag_idx, pagina in enumerate(pdf.pages, 1):
        texto = pagina.extract_text()
        
        # Procurar "Seq 10" ou "10 125" (formato do bloco)
        if "Seq. NIT Código" in texto and "10 125" in texto:
            print(f"📄 PÁGINA {pag_idx}: Seq 10 ENCONTRADA!")
            print()
            
            # Extrair blocos
            blocos = _extrair_blocos_vinculos(texto)
            print(f"Blocos encontrados nesta página: {len(blocos)}")
            
            for i, bloco in enumerate(blocos, 1):
                if bloco['seq'] == '10':
                    print(f"\n✅ Bloco {i} é Seq 10:")
                    print(f"   NIT: {bloco['nit']}")
                    print(f"   CNPJ: {bloco['cnpj']}")
                    print(f"   Tem 'Remunerações': {'Remunerações' in bloco['texto']}")
                    print(f"   Tem 'Competência': {'Competência' in bloco['texto']}")
                    
                    # Mostrar últimas 10 linhas do bloco
                    linhas = bloco['texto'].split('\n')
                    print(f"\n   Últimas 10 linhas do bloco:")
                    for linha in linhas[-10:]:
                        print(f"      {linha[:70]}")
            
            print()
            print("─" * 70)
            
            # Ver se tem rodapé
            if "O INSS poderá rever" in texto:
                print("⚠️  TEM RODAPÉ nesta página - bloco pode estar cortado")
            else:
                print("✅ Sem rodapé - bloco completo")
            
            print()
            print("─" * 70)
        
        # Procurar seção de Remunerações órfã (continuação)
        zona_util = _extrair_zona_util(texto)
        if "Competência" in zona_util[:500] and "Remuneração" in zona_util[:500]:
            if not "Matrícula do Tipo Filiado" in zona_util[:300]:
                print(f"📄 PÁGINA {pag_idx}: Possível SEÇÃO ÓRFÃ de Remunerações")
                linhas = zona_util.split('\n')[:15]
                print("   Primeiras 15 linhas da zona útil:")
                for linha in linhas:
                    print(f"      {linha[:70]}")
                print()
