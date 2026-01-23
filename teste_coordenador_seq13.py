from extrator.tipos.coordenador_remuneracoes import extrair_remuneracoes_coordenado

pdf_path = "cnis/CNIS_JOAO_CARLOS.pdf"

print("🔍 TESTANDO COORDENADOR COMPLETO")
print("="*70)

registros = extrair_remuneracoes_coordenado(pdf_path)

# Filtrar Seq 13
seq13 = [r for r in registros if r['seq'] == '13']

print(f"\n📊 SEQ 13: {len(seq13)} registros")
for r in seq13:
    print(f"  Página {r['pagina']}: {r['competencia']} - {r['remuneracao']} - {r['tipo_vinculo']}")

if len(seq13) == 4:
    print("\n✅ TODOS OS 4 REGISTROS DA SEQ 13 CAPTURADOS!")
else:
    print(f"\n⚠️  ESPERADO 4, CAPTURADO {len(seq13)}")
    esperados = ['08/2025', '09/2025', '10/2025', '11/2025']
    capturados = [r['competencia'] for r in seq13]
    faltantes = [e for e in esperados if e not in capturados]
    if faltantes:
        print(f"   Faltantes: {faltantes}")

# Ver total
print(f"\n📊 TOTAL GERAL: {len(registros)} registros")
