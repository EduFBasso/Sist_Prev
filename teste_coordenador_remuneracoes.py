"""
Teste do Coordenador de Remunerações

VALIDA:
    1. Extração delegada por tipo (CLT + FACULTATIVO)
    2. Baseline preservado: 178 remunerações
    3. Distribuição correta: 163 CLT + 15 FACULTATIVO
    4. Campos completos em cada registro
"""

from extrator.tipos.coordenador_remuneracoes import (
    extrair_remuneracoes_coordenado,
    validar_baseline_remuneracoes
)

PDF_TESTE = "cnis/CNIS_JOAO_CARLOS.pdf"

def main():
    print("=" * 70)
    print("🔍 TESTE: COORDENADOR DE REMUNERAÇÕES")
    print("=" * 70)
    print()
    
    print(f"📄 Arquivo: {PDF_TESTE}")
    print()
    
    print("─" * 70)
    print("1️⃣  EXTRAÇÃO COORDENADA (DELEGAÇÃO AUTOMÁTICA)")
    print("─" * 70)
    print()
    
    print("🔄 Processando PDF...")
    print("   • Detectando tipos de vínculos automaticamente")
    print("   • Delegando para processadores especializados")
    print("   • CLT → processar_remuneracoes_clt() (3 campos)")
    print("   • FACULTATIVO → processar_contribuicoes_facultativo() (5 campos)")
    print()
    
    registros = extrair_remuneracoes_coordenado(PDF_TESTE)
    
    print(f"✅ Extraídas: {len(registros)} remunerações")
    print()
    
    print("─" * 70)
    print("2️⃣  VALIDAÇÃO DE BASELINE")
    print("─" * 70)
    print()
    
    validacao = validar_baseline_remuneracoes(registros)
    
    print(f"Total extraído: {validacao['total']}")
    print(f"Distribuição por tipo:")
    for tipo, count in validacao['por_tipo'].items():
        emoji = "🏢" if tipo == "CLT" else "👤" if tipo == "FACULTATIVO" else "❓"
        print(f"  {emoji} {tipo:15s}: {count:3d} remunerações")
    print()
    print(validacao['mensagem'])
    print()
    
    print("─" * 70)
    print("3️⃣  AMOSTRAS DE REGISTROS")
    print("─" * 70)
    print()
    
    # Separar por tipo
    clt_registros = [r for r in registros if r.get('tipo_vinculo') == 'CLT']
    fac_registros = [r for r in registros if r.get('tipo_vinculo') == 'FACULTATIVO']
    
    # Mostrar 3 primeiros CLT
    if clt_registros:
        print("🏢 AMOSTRAS CLT (3 primeiros):")
        for i, reg in enumerate(clt_registros[:3], 1):
            print(f"   {i}. Seq {reg.get('seq')}, CNPJ {reg.get('cnpj')}, "
                  f"{reg.get('competencia')} = R$ {reg.get('remuneracao')}")
        print()
    
    # Mostrar 3 primeiros FACULTATIVO
    if fac_registros:
        print("👤 AMOSTRAS FACULTATIVO (3 primeiros):")
        for i, reg in enumerate(fac_registros[:3], 1):
            print(f"   {i}. Seq {reg.get('seq')}, "
                  f"{reg.get('competencia')} = R$ {reg.get('remuneracao')}")
        print()
    
    print("─" * 70)
    print("4️⃣  VALIDAÇÃO DE CAMPOS")
    print("─" * 70)
    print()
    
    # Validar que todos os registros têm campos obrigatórios
    campos_obrigatorios = ['seq', 'competencia', 'remuneracao', 'pagina', 'tipo_vinculo']
    registros_completos = 0
    
    for reg in registros:
        if all(campo in reg for campo in campos_obrigatorios):
            registros_completos += 1
    
    percentual = (registros_completos / len(registros) * 100) if registros else 0
    
    print(f"Campos obrigatórios: {', '.join(campos_obrigatorios)}")
    print(f"Registros completos: {registros_completos}/{len(registros)} ({percentual:.1f}%)")
    print()
    
    if percentual == 100:
        print("✅ Todos os registros têm campos obrigatórios")
    else:
        print("⚠️  Alguns registros estão incompletos")
    print()
    
    print("─" * 70)
    print("5️⃣  ESTATÍSTICAS POR SEQUÊNCIA")
    print("─" * 70)
    print()
    
    # Contar por sequência
    por_seq = {}
    for reg in registros:
        seq = reg.get('seq', '?')
        por_seq[seq] = por_seq.get(seq, 0) + 1
    
    # Ordenar por sequência
    for seq in sorted(por_seq.keys(), key=lambda x: int(x) if x.isdigit() else 999):
        count = por_seq[seq]
        # Identificar tipo pela sequência
        tipo_emoji = "🏢" if int(seq) <= 10 else "👤" if int(seq) <= 13 else "❓"
        print(f"  Seq {seq:2s}: {count:3d} remunerações {tipo_emoji}")
    print()
    
    print("─" * 70)
    print("6️⃣  RESUMO FINAL")
    print("─" * 70)
    print()
    
    if validacao['valido'] and registros_completos == len(registros):
        print("✅ COORDENADOR FUNCIONANDO PERFEITAMENTE!")
        print()
        print("Responsabilidade 3: ✅ Coordenar Remunerações")
        print(f"   • Total extraído: {len(registros)} remunerações")
        print(f"   • CLT: {len(clt_registros)} (delegado para processar_remuneracoes_clt)")
        print(f"   • FACULTATIVO: {len(fac_registros)} (delegado para processar_contribuicoes_facultativo)")
        print("   • Baseline: ✅ 178 remunerações preservadas")
        print("   • Campos: ✅ Todos completos")
        print("   • Detecção automática: ✅ Funcionando")
        print()
        print("🎯 PRÓXIMO PASSO: Integrar coordenador no CLI")
        print("   • Substituir extrair_remuneracoes_texto() monolítico (560 linhas)")
        print("   • Usar processar_cnis_completo() do orquestrador")
        print("   • Reduzir CLI de 922 para ~150 linhas")
    else:
        print("❌ VALIDAÇÃO FALHOU!")
        if not validacao['valido']:
            print(f"   • Baseline incorreto: {validacao['total']} (esperado 178)")
        if registros_completos != len(registros):
            print(f"   • Registros incompletos: {len(registros) - registros_completos}")
    
    print()
    print("=" * 70)

if __name__ == "__main__":
    main()
