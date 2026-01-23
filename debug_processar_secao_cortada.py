import pdfplumber
import re
from extrator.tipos.coordenador_remuneracoes import _extrair_zona_util
from extrator.tipos.facultativo import processar_contribuicoes_facultativo

pdf_path = "cnis/CNIS_JOAO_CARLOS.pdf"

with pdfplumber.open(pdf_path) as pdf:
    pag7 = pdf.pages[6]  # página 7
    texto7 = pag7.extract_text()
    zona_util = _extrair_zona_util(texto7)
    
    print("🔍 SIMULANDO _processar_secao_cortada() PARA SEQ 11")
    print("="*70)
    
    # Simular bloco_cortado
    bloco_cortado = {
        'seq': '11',
        'tipo': 'FACULTATIVO'
    }
    
    seq = bloco_cortado['seq']
    tipo = bloco_cortado.get('tipo', 'CLT')
    
    print(f"Seq: {seq}, Tipo: {tipo}")
    print()
    
    # Buscar seção FACULTATIVO
    match_secao = re.search(r'Contribuições\s+Competência\s+Data\s+Pgto', 
                           zona_util, re.IGNORECASE)
    
    print(f"Match 'Contribuições Competência Data Pgto': {match_secao is not None}")
    
    if match_secao:
        inicio_secao = match_secao.start()
        print(f"Início da seção: {inicio_secao}")
        
        # Extrair até o próximo bloco
        match_proximo_bloco = re.search(r'Matrícula do Tipo Filiado', zona_util[inicio_secao:], re.IGNORECASE)
        print(f"Match 'Matrícula do Tipo Filiado': {match_proximo_bloco is not None}")
        
        # Tentar outro marcador
        if not match_proximo_bloco:
            match_proximo_bloco = re.search(r'Seq\.\s+NIT\s+Origem do Vínculo', zona_util[inicio_secao:], re.IGNORECASE)
            print(f"Match 'Seq. NIT Origem do Vínculo': {match_proximo_bloco is not None}")
        
        if match_proximo_bloco:
            secao = zona_util[inicio_secao:inicio_secao + match_proximo_bloco.start()]
        else:
            secao = zona_util[inicio_secao:]
        
        print(f"\n📝 SEÇÃO EXTRAÍDA ({len(secao)} chars):")
        print("="*70)
        print(secao[:500])
        print()
        
        # Processar
        registros = []
        processar_contribuicoes_facultativo(secao, seq, 7, registros)
        
        print(f"\n📊 REGISTROS: {len(registros)}")
        for r in registros:
            print(f"  {r['competencia']} - {r['remuneracao']}")
