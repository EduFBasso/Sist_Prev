import pdfplumber
import re

with pdfplumber.open("CNIS_JOAO_CARLOS.pdf") as pdf:
    texto = pdf.pages[0].extract_text()
    
    # Encontrar todos os blocos "Matrícula do Tipo Filiado"
    pos = 0
    bloco_num = 0
    while True:
        inicio = texto.find("Matrícula do Tipo Filiado", pos)
        if inicio == -1:
            break
        
        proximo = texto.find("Matrícula do Tipo Filiado", inicio + 26)
        fim = proximo if proximo != -1 else len(texto)
        
        bloco = texto[inicio:fim]
        
        # Extrair Seq
        seq_match = re.search(r"Seq\.\s+.*?\n(\d+)", bloco)
        seq = seq_match.group(1) if seq_match else "?"
        
        print(f"\n{'='*60}")
        print(f"BLOCO {bloco_num + 1} - Seq {seq}")
        print('='*60)
        print(bloco[:500])
        
        # Procurar remunerações
        if "Remunerações" in bloco:
            print("\n>>> TEM SEÇÃO REMUNERAÇÕES")
            inicio_remun = bloco.find("Remunerações")
            secao = bloco[inicio_remun:]
            linhas = secao.split('\n')[:10]
            for i, l in enumerate(linhas):
                print(f"  {i}: {l[:80]}")
        
        bloco_num += 1
        pos = fim
