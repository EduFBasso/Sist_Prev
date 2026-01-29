"""
Análise de Índices INPC vs SELIC

Compara os fatores de correção para entender a diferença de escala.
"""

import csv

# Carregar INPC
inpc = {}
with open('fat_inpc_selic/inpc_fatores.csv', 'r', encoding='utf-8') as f:
    reader = csv.DictReader(f, delimiter=';')
    for row in reader:
        comp = row['Competencia']
        fator = float(row['Fator'])
        inpc[comp] = fator

# Carregar SELIC
selic = {}
with open('fat_inpc_selic/selic_fatores.csv', 'r', encoding='utf-8') as f:
    reader = csv.DictReader(f, delimiter=';')
    for row in reader:
        comp = row['Competencia']
        fator = float(row['Fator'])
        selic[comp] = fator

print("\n" + "="*80)
print("   📊 ANÁLISE DE ÍNDICES - INPC vs SELIC")
print("="*80 + "\n")

# Períodos cobertos
print(f"INPC: {len(inpc)} competências")
print(f"SELIC: {len(selic)} competências\n")

# Comparar competências em comum (1995-2024)
print("📋 AMOSTRA: Janeiro de cada ano (1995-2024)")
print("-" * 80)
print(f"{'Ano':<8} {'INPC':>15} {'SELIC':>15} {'Razão':>15} {'Remuneração':>15}")
print("-" * 80)

valor_base = 286.25  # Salário exemplo do João Carlos

for ano in range(1995, 2025):
    comp = f"01/{ano}"
    if comp in inpc and comp in selic:
        i = inpc[comp]
        s = selic[comp]
        razao = s / i if i > 0 else 0
        
        valor_inpc = valor_base * i
        valor_selic = valor_base * s
        
        print(f"{ano:<8} {i:>15.6f} {s:>15.6f} {razao:>15.2f}x   R$ {valor_inpc:>12,.2f} | R$ {valor_selic:>12,.2f}")

print("-" * 80)

# Último índice disponível
ultimo_inpc = max(inpc.keys())
ultimo_selic = max(selic.keys())

print(f"\n📅 Última competência:")
print(f"   INPC:  {ultimo_inpc} = {inpc[ultimo_inpc]:.6f}")
print(f"   SELIC: {ultimo_selic} = {selic[ultimo_selic]:.6f}")

# Análise: Se SELIC/INPC ≈ constante, são escalas diferentes
# Se SELIC/INPC varia muito, são metodologias diferentes

print("\n" + "="*80)
print("   💡 INTERPRETAÇÃO:")
print("="*80)
print("""
Se a RAZÃO (SELIC/INPC) for aproximadamente CONSTANTE ao longo dos anos:
→ Os fatores estão em ESCALAS DIFERENTES (datas-base diferentes)
→ Solução: Normalizar ambos para mesma data-base

Se a RAZÃO VARIAR MUITO:
→ São METODOLOGIAS DIFERENTES de cálculo
→ SELIC acumula juros compostos mais agressivamente
→ Ambos estão corretos, mas representam índices de natureza diferente

ATENÇÃO: No sistema VBA original, SELIC NÃO estava implementado.
Os fatores SELIC foram gerados posteriormente e precisam VALIDAÇÃO JURÍDICA.
""")
