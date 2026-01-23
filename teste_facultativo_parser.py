from extrator.tipos.facultativo import processar_contribuicoes_facultativo

# Simular seção de contribuições órfã (Seq 11)
secao_contribuicoes = """Contribuições
Competência Data Pgto. Contribuição Salário Contribuição Indicadores Competência Data Pgto. Contribuição Salário Contribuição Indicadores
09/2019 18/09/2019 200,00 1.000,00 PREC-FACULTCONC 10/2019 22/11/2019 199,60 998,00 PREC-FACULTCONC"""

registros = []

print("🔍 TESTANDO processar_contribuicoes_facultativo()")
print("="*70)
print(f"\nTexto da seção:")
print(secao_contribuicoes)
print()

processar_contribuicoes_facultativo(secao_contribuicoes, seq='11', pagina_idx=7, registros=registros)

print(f"📊 REGISTROS CAPTURADOS: {len(registros)}")
for r in registros:
    print(f"  Seq {r['seq']}: {r['competencia']} - {r['remuneracao']} - {r['indicadores']}")
