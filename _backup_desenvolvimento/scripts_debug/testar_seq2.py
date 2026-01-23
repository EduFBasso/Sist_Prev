import csv

with open('cnis/JOAO_CARLOS_EDUARDO_FIGUEIREDO_BASSO/cnis_remuneracoes.csv', encoding='utf-8') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
    
print(f"Total: {len(rows)} registros")

seq2 = [r for r in rows if r['Seq'] == '2']
print(f"\nSeq 2: {len(seq2)} registros")

comps = sorted(set([r['Competencia'] for r in seq2]))
print(f"\nCompetências da Seq 2:")
for c in comps:
    print(f"  {c}")

print(f"\nTem 01/1998? {'01/1998' in comps}")
print(f"Tem 11/1998? {'11/1998' in comps}")
