from extrator.tipos.facultativo import processar_contribuicoes_facultativo

# Texto real da Seq 13
secao_seq13 = """Contribuições
Competência Data Pgto. Contribuição Salário Contribuição Indicadores Competência Data Pgto. Contribuição Salário Contribuição Indicadores                                                                                                    08/2025 15/09/2025 303,60 1.518,00 PREC-FACULTCONC 09/2025 15/09/2025 303,60 1.518,00 PREC-FACULTCONC
10/2025 14/11/2025 303,60 1.518,00 PREC-FACULTCONC 11/2025 15/12/2025 303,60 1.518,00 PREC-FACULTCONC
"""

registros = []

print("🔍 TESTANDO PARSER FACULTATIVO NA SEQ 13")
print("="*70)
print("\nTexto da seção:")
print(secao_seq13)
print("\n" + "="*70)

processar_contribuicoes_facultativo(secao_seq13, seq='13', pagina_idx=7, registros=registros)

print(f"\n📊 REGISTROS CAPTURADOS: {len(registros)}")
for r in registros:
    print(f"  {r['competencia']} - {r['remuneracao']} - {r['indicadores']}")

if len(registros) == 4:
    print("\n✅ TODOS OS 4 REGISTROS CAPTURADOS!")
else:
    print(f"\n⚠️  ESPERADO 4, CAPTURADO {len(registros)}")
    esperados = ['08/2025', '09/2025', '10/2025', '11/2025']
    capturados = [r['competencia'] for r in registros]
    faltantes = [e for e in esperados if e not in capturados]
    if faltantes:
        print(f"   Faltantes: {faltantes}")
