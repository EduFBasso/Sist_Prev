"""Testar processamento da Seq 10 especificamente"""

import pdfplumber
from extrator.tipos.coordenador_remuneracoes import extrair_remuneracoes_coordenado

PDF_TESTE = "cnis/CNIS_JOAO_CARLOS.pdf"

print("🔍 TESTANDO EXTRAÇÃO COM DEBUG")
print("=" * 70)
print()

registros = extrair_remuneracoes_coordenado(PDF_TESTE)

print()
print("─" * 70)
print("RESULTADO:")
print("─" * 70)

# Contar por Seq
por_seq = {}
for reg in registros:
    seq = reg.get('seq', '?')
    por_seq[seq] = por_seq.get(seq, 0) + 1

for seq in sorted(por_seq.keys(), key=lambda x: int(x) if x.isdigit() else 999):
    count = por_seq[seq]
    esperado = {
        '1': 5, '2': 31, '3': 2, '4': 6, '5': 3, '6': 4, '7': 7,
        '8': 38, '9': 21, '10': 46, '11': 2, '12': 9, '13': 4
    }.get(seq, 0)
    
    status = "✅" if count == esperado else f"⚠️ (faltam {esperado - count})"
    print(f"Seq {seq:2s}: {count:2d}/{esperado:2d} {status}")

print()
print(f"Total: {len(registros)}/178")
