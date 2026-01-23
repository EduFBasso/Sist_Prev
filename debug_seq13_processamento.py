import pdfplumber
from extrator.tipos.coordenador_remuneracoes import _extrair_zona_util, _extrair_blocos_vinculos
from extrator.tipos.facultativo import processar_contribuicoes_facultativo

pdf_path = "cnis/CNIS_JOAO_CARLOS.pdf"

with pdfplumber.open(pdf_path) as pdf:
    pag7 = pdf.pages[6]
    texto7 = pag7.extract_text()
    zona_util = _extrair_zona_util(texto7)
    blocos = _extrair_blocos_vinculos(zona_util)
    
    for bloco in blocos:
        if bloco['seq'] == '13':
            print("🔍 PROCESSANDO BLOCO SEQ 13")
            print("="*70)
            
            # Extrair seção Contribuições
            import re
            match_contrib = re.search(r'Contribuições\s*(.*?)(?:Remunerações|Matrícula|$)', 
                                     bloco['texto'], re.DOTALL | re.IGNORECASE)
            
            if match_contrib:
                secao_contrib = match_contrib.group(1)
                
                print(f"\n📝 SEÇÃO CONTRIBUIÇÕES ({len(secao_contrib)} chars):")
                print(secao_contrib)
                print("\n" + "="*70)
                
                # Processar
                registros = []
                processar_contribuicoes_facultativo(secao_contrib, seq='13', pagina_idx=7, registros=registros)
                
                print(f"\n📊 REGISTROS PROCESSADOS: {len(registros)}")
                for r in registros:
                    print(f"  {r['competencia']} - {r['remuneracao']}")
                
                # Verificar se tem "Seq." na seção que poderia parar o processamento
                if "Seq." in secao_contrib:
                    print(f"\n⚠️  TEM 'Seq.' NA SEÇÃO - pode estar parando antes!")
                    linhas = secao_contrib.split('\n')
                    for i, linha in enumerate(linhas):
                        if 'Seq.' in linha:
                            print(f"   Linha {i}: {linha}")
            else:
                print("❌ Não encontrou seção Contribuições")
