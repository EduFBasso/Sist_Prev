import json
from urllib.request import urlopen, Request

# Série 4390 - últimos meses
url = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.4390/dados?formato=json&dataInicial=01/01/2024&dataFinal=22/01/2026"
req = Request(url, headers={'User-Agent': 'Mozilla/5.0'})
data = json.loads(urlopen(req).read())

print("=== Série 4390 - Últimos 15 registros ===")
for item in data[-15:]:
    print(f"{item['data']}: {item['valor']}")

# Série 11 - Meta SELIC (para comparação)
url2 = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.11/dados?formato=json&dataInicial=01/01/2024&dataFinal=22/01/2026"
req2 = Request(url2, headers={'User-Agent': 'Mozilla/5.0'})
data2 = json.loads(urlopen(req2).read())

print("\n=== Série 11 - Meta SELIC (% a.a.) - Últimos 10 ===")
for item in data2[-10:]:
    print(f"{item['data']}: {item['valor']}% ao ano")

# Calcular taxa mensal esperada
if data2:
    taxa_anual = float(data2[-1]['valor'])
    taxa_mensal = (pow(1 + taxa_anual/100, 1/12) - 1) * 100
    print(f"\n✓ Taxa SELIC anual atual: {taxa_anual}%")
    print(f"✓ Taxa mensal equivalente: {taxa_mensal:.4f}%")
    print(f"\n⚠️  Série 4390 mostra: {data[-1]['valor']}")
    print(f"⚠️  Se for % mensal, deveria ser ~{taxa_mensal:.2f}, não {data[-1]['valor']}!")
