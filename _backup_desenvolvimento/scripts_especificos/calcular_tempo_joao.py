#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Calcula Tempo de Contribuição e Cenários de Aposentadoria
"""

import csv
from pathlib import Path
from datetime import datetime

def detectar_encoding(arquivo):
    """Detecta encoding do arquivo"""
    try:
        with open(arquivo, 'r', encoding='utf-8') as f:
            f.read()
        return 'utf-8'
    except UnicodeDecodeError:
        return 'cp1252'

def calcular_tempo_contribuicao():
    """Calcula tempo total de contribuição do CNIS"""
    
    print("\n" + "="*80)
    print("⏱️  CÁLCULO DE TEMPO DE CONTRIBUIÇÃO - JOÃO CARLOS")
    print("="*80)
    
    # Localizar pasta do cliente
    pasta_cnis = Path("cnis")
    pastas_cliente = [p for p in pasta_cnis.iterdir() if p.is_dir()]
    
    if not pastas_cliente:
        print("\n❌ Nenhuma pasta de cliente encontrada")
        return
    
    pasta_cliente = pastas_cliente[0]
    print(f"\n📁 Cliente: {pasta_cliente.name}")
    
    # Ler vínculos
    arquivo_vinculos = list(pasta_cliente.glob("*vinculos_estruturado.csv"))[0]
    encoding = detectar_encoding(arquivo_vinculos)
    
    vinculos = []
    with open(arquivo_vinculos, 'r', encoding=encoding) as f:
        reader = csv.DictReader(f, delimiter=';')
        for row in reader:
            vinculos.append(row)
    
    print(f"\n📊 Total de vínculos: {len(vinculos)}")
    print("\n" + "─"*80)
    print(f"{'Seq':>3} │ {'Início':12} │ {'Fim':12} │ {'Dias':>8} │ {'Empresa':40}")
    print("─"*80)
    
    total_dias = 0
    primeira_data = None
    ultima_data = None
    
    for v in vinculos:
        seq = v.get('Seq', '')
        inicio_str = v.get('DataInicio', '')
        fim_str = v.get('DataFim', '')
        empresa = v.get('Empresa', '')[:38]
        
        # Calcular dias
        dias = 0
        if inicio_str and fim_str:
            try:
                # Tentar vários formatos de data
                for fmt in ['%d/%m/%Y', '%Y-%m-%d']:
                    try:
                        inicio = datetime.strptime(inicio_str, fmt)
                        fim = datetime.strptime(fim_str, fmt)
                        dias = (fim - inicio).days + 1
                        
                        if not primeira_data or inicio < primeira_data:
                            primeira_data = inicio
                        if not ultima_data or fim > ultima_data:
                            ultima_data = fim
                        
                        break
                    except:
                        continue
            except:
                pass
        
        total_dias += dias
        
        print(f"{seq:>3}  │ {inicio_str:12} │ {fim_str:12} │ {dias:>8} │ {empresa}")
    
    print("─"*80)
    
    # Cálculos
    total_anos = total_dias / 365
    total_meses = total_dias / 30.44
    
    print(f"\n📈 RESUMO DO TEMPO:")
    print(f"   • Total em dias: {total_dias:,} dias")
    print(f"   • Total em meses: {total_meses:.1f} meses")
    print(f"   • Total em anos: {total_anos:.2f} anos")
    
    if primeira_data and ultima_data:
        print(f"\n📅 PERÍODO:")
        print(f"   • Primeira contribuição: {primeira_data.strftime('%d/%m/%Y')}")
        print(f"   • Última contribuição: {ultima_data.strftime('%d/%m/%Y')}")
    
    # Idade atual
    data_nascimento = datetime(1973, 3, 22)
    hoje = datetime.now()
    idade_atual = (hoje - data_nascimento).days / 365
    
    print(f"\n👤 DADOS PESSOAIS:")
    print(f"   • Data de nascimento: 22/03/1973")
    print(f"   • Idade atual: {idade_atual:.1f} anos")
    
    # Cenários de aposentadoria
    print(f"\n" + "="*80)
    print(f"🎯 CENÁRIOS DE APOSENTADORIA")
    print("="*80)
    
    # Cenário 1: Hoje (se tiver idade)
    anos_faltam_idade = max(0, 65 - idade_atual)
    coef_atual = 0.60 + 0.02 * max(0, total_anos - 20)
    
    print(f"\n📊 CENÁRIO 1: Aposentar HOJE (idade {idade_atual:.1f} anos)")
    if idade_atual >= 65:
        print(f"   ✅ Idade mínima: OK (65 anos)")
        print(f"   ✅ Tempo: {total_anos:.2f} anos")
        print(f"   💰 Coeficiente: {coef_atual:.2%}")
        print(f"   📌 Status: PODE SE APOSENTAR")
    else:
        print(f"   ❌ Faltam {anos_faltam_idade:.1f} anos para idade mínima (65 anos)")
        print(f"   ✅ Tempo: {total_anos:.2f} anos")
        print(f"   💰 Coeficiente se aposentasse: {coef_atual:.2%}")
    
    # Cenário 2: Aos 65 anos (com tempo atual)
    print(f"\n📊 CENÁRIO 2: Aos 65 anos (SEM contribuir mais)")
    print(f"   ✅ Idade: 65 anos (em {anos_faltam_idade:.1f} anos)")
    print(f"   ✅ Tempo: {total_anos:.2f} anos")
    print(f"   💰 Coeficiente: {coef_atual:.2%}")
    print(f"   📌 Valor: Média × {coef_atual:.2%}")
    
    # Cenário 3: Contribuir até 100%
    anos_para_100 = max(0, 40 - total_anos)
    idade_100 = idade_atual + anos_para_100
    
    print(f"\n📊 CENÁRIO 3: Contribuir até coeficiente 100%")
    print(f"   ⏱️  Faltam: {anos_para_100:.1f} anos de contribuição")
    print(f"   📅 Idade quando atingir: {idade_100:.1f} anos")
    print(f"   💰 Coeficiente: 100%")
    print(f"   📌 Valor: Média × 100% (MÁXIMO)")
    
    # Tabela de evolução
    print(f"\n" + "="*80)
    print(f"📈 EVOLUÇÃO DO COEFICIENTE")
    print("="*80)
    print(f"{'Anos':>6} │ {'Coeficiente':>12} │ {'Exemplo R$ 2.500':>18}")
    print("─"*80)
    
    for anos in [15, 20, 25, 30, 35, 40]:
        coef = 0.60 + 0.02 * max(0, anos - 20)
        valor = 2500 * coef
        if anos <= total_anos:
            print(f"{anos:>6} │ {coef:>11.2%} │ R$ {valor:>13,.2f}  ✅ JÁ TEM")
        else:
            falta = anos - total_anos
            print(f"{anos:>6} │ {coef:>11.2%} │ R$ {valor:>13,.2f}  (faltam {falta:.1f} anos)")
    
    print("\n" + "="*80)
    print(f"💡 RECOMENDAÇÃO:")
    
    if total_anos < 20:
        print(f"   • Continue contribuindo! Você tem {total_anos:.1f} anos")
        print(f"   • Ao atingir 20 anos, já terá coeficiente de 60%")
        print(f"   • Cada ano adicional = +2% no valor")
    elif total_anos < 35:
        print(f"   • Você tem {total_anos:.1f} anos (coeficiente {coef_atual:.1%})")
        print(f"   • Contribuindo mais {35-total_anos:.1f} anos → 90%")
        print(f"   • Contribuindo mais {40-total_anos:.1f} anos → 100% (máximo)")
    else:
        print(f"   • Você tem {total_anos:.1f} anos (coeficiente {coef_atual:.1%})")
        print(f"   • Está perto do máximo!")
        print(f"   • Faltam {40-total_anos:.1f} anos para 100%")
    
    print("="*80 + "\n")
    
    return total_anos, coef_atual

if __name__ == "__main__":
    calcular_tempo_contribuicao()
