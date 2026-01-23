import pdfplumber
from extrator.tipos.coordenador_remuneracoes import extrair_remuneracoes_coordenado

pdf_path = "cnis/CNIS_JOAO_CARLOS.pdf"

# Executar coordenador
registros = extrair_remuneracoes_coordenado(pdf_path)

# Contar por tipo
clt = [r for r in registros if r.get('tipo_vinculo') == 'CLT']
facultativo = [r for r in registros if r.get('tipo_vinculo') == 'FACULTATIVO']

print(f"📊 TIPOS DE VÍNCULOS:")
print(f"   CLT: {len(clt)}")
print(f"   FACULTATIVO: {len(facultativo)}")
print(f"   TOTAL: {len(registros)}")

# Ver quais sequências facultativas temos
if facultativo:
    print(f"\n📋 FACULTATIVOS CAPTURADOS:")
    for r in facultativo:
        print(f"   Seq {r['seq']}: {r['competencia']} - {r['remuneracao']}")
else:
    print(f"\n⚠️  NENHUM FACULTATIVO CAPTURADO!")

# Procurar na página 7
print(f"\n🔍 PROCURANDO FACULTATIVOS NA PÁGINA 7...")
with pdfplumber.open(pdf_path) as pdf:
    pagina7 = pdf.pages[6]  # página 7 (índice 6)
    texto = pagina7.extract_text()
    
    # Procurar "Seq 11", "Seq 12", "Seq 13"
    for seq in ['11', '12', '13']:
        if f"NIT 125.37781.66-{seq}" in texto or f" {seq} 125" in texto:
            print(f"   ✅ Seq {seq} ENCONTRADA na página 7")
            # Pegar contexto
            linhas = texto.split('\n')
            for i, linha in enumerate(linhas):
                if f" {seq} 125" in linha or f"66-{seq}" in linha:
                    print(f"      Contexto: {linha}")
                    if i+1 < len(linhas):
                        print(f"               {linhas[i+1]}")
                    break
