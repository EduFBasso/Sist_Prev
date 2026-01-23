"""Ver total de blocos"""
import pdfplumber
from extrator.tipos.coordenador_remuneracoes import _extrair_blocos_vinculos

pdf = pdfplumber.open('cnis/CNIS_JOAO_CARLOS.pdf')
total = 0
for i, p in enumerate(pdf.pages, 1):
    texto = p.extract_text()
    blocos = _extrair_blocos_vinculos(texto)
    total += len(blocos)
    print(f"Página {i}: {len(blocos)} blocos")

print(f"\nTotal: {total} blocos")
