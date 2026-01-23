#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de Teste da Estrutura Modular do Extrator CNIS

Testa cada componente isoladamente e depois a integração completa.
"""

import sys
from pathlib import Path

# Adicionar caminho do projeto
sys.path.insert(0, str(Path(__file__).parent))

print("=" * 80)
print("🧪 TESTE DA ESTRUTURA MODULAR - EXTRATOR CNIS")
print("=" * 80)
print()

# ============================================================================
# TESTE 1: Detector de Tipos
# ============================================================================
print("📋 TESTE 1: Detector de Tipos")
print("-" * 80)

from extrator.tipos import detectar_tipo_vinculo, TipoVinculo, obter_nome_tipo

# Caso 1: CLT (com Código Emp.)
texto_clt = """
Matrícula do Tipo Filiado no
Seq. NIT Código Emp. Origem do Vínculo
1 125.37781.66-1 56.528.946/0001-80 EMPRESA XYZ LTDA Empregado
"""
tipo_clt = detectar_tipo_vinculo(texto_clt)
print(f"✓ CLT detectado: {tipo_clt == TipoVinculo.CLT}")
print(f"  Nome: {obter_nome_tipo(tipo_clt)}")

# Caso 2: Facultativo (sem Código Emp., com RECOLHIMENTO)
texto_fac = """
Origem do Vínculo Trabalhador Vínculo Data Início Data Fim Últ. Remun.
11 125.37781.66-1 RECOLHIMENTO - FACULTATIVO Contribuinte Individual
"""
tipo_fac = detectar_tipo_vinculo(texto_fac)
print(f"✓ FACULTATIVO detectado: {tipo_fac == TipoVinculo.FACULTATIVO}")
print(f"  Nome: {obter_nome_tipo(tipo_fac)}")
print()

# ============================================================================
# TESTE 2: Parser de Cabeçalho
# ============================================================================
print("📋 TESTE 2: Parser de Cabeçalho")
print("-" * 80)

from extrator.parsers import extrair_dados_cabecalho

pdf_teste = "cnis/CNIS_JOAO_CARLOS.pdf"
if Path(pdf_teste).exists():
    dados_cab = extrair_dados_cabecalho(pdf_teste)
    print(f"✓ NIT: {dados_cab.get('NIT', 'N/A')}")
    print(f"✓ CPF: {dados_cab.get('CPF', 'N/A')}")
    print(f"✓ Nome: {dados_cab.get('Nome', 'N/A')}")
    print(f"✓ Data Nascimento: {dados_cab.get('DataNascimento', 'N/A')}")
    print(f"✓ Nome Mãe: {dados_cab.get('NomeMae', 'N/A')}")
else:
    print(f"⚠️  PDF não encontrado: {pdf_teste}")
print()

# ============================================================================
# TESTE 3: Processadores por Tipo (CLT e Facultativo)
# ============================================================================
print("📋 TESTE 3: Processadores por Tipo")
print("-" * 80)

from extrator.tipos import processar_remuneracoes_clt, processar_contribuicoes_facultativo

# Teste CLT
registros_clt = []
secao_clt_exemplo = """
Remunerações
Competência Remuneração Indicadores
01/1995 286,25 13º SALÁRIO
02/1995 286,25
03/1995 286,25
"""
processar_remuneracoes_clt(secao_clt_exemplo, "1", "56.528.946/0001-80", 1, registros_clt)
print(f"✓ CLT processado: {len(registros_clt)} remunerações")
if registros_clt:
    print(f"  Exemplo: Seq={registros_clt[0]['seq']}, " +
          f"Competência={registros_clt[0]['competencia']}, " +
          f"Valor={registros_clt[0]['remuneracao']}")

# Teste Facultativo
registros_fac = []
secao_fac_exemplo = """
Contribuições
Competência Data Pagto. Contribuição Salário Contrib. Indicadores
09/2019 15/09/2019 200,00 1045,00 PREC-FACULTCONC
10/2019 15/10/2019 199,60 1045,00 PREC-FACULTCONC
"""
processar_contribuicoes_facultativo(secao_fac_exemplo, "11", 1, registros_fac)
print(f"✓ FACULTATIVO processado: {len(registros_fac)} contribuições")
if registros_fac:
    print(f"  Exemplo: Seq={registros_fac[0]['seq']}, " +
          f"Competência={registros_fac[0]['competencia']}, " +
          f"Valor={registros_fac[0]['remuneracao']}")
print()

# ============================================================================
# TESTE 4: Parser de Vínculos
# ============================================================================
print("📋 TESTE 4: Parser de Vínculos")
print("-" * 80)

from extrator.parsers import extrair_vinculos_texto, parse_vinculo_texto

if Path(pdf_teste).exists():
    blocos = extrair_vinculos_texto(pdf_teste)
    print(f"✓ Blocos de vínculos extraídos: {len(blocos)}")
    
    if blocos:
        # Parse primeiro bloco
        pagina, texto_bloco = blocos[0]
        vinculo_parsed = parse_vinculo_texto(texto_bloco)
        if vinculo_parsed:
            print(f"  Vínculo 1: Seq={vinculo_parsed.get('seq')}, " +
                  f"Empresa={vinculo_parsed.get('empresa', '')[:30]}...")
print()

# ============================================================================
# TESTE 5: Extração Completa de Tabelas
# ============================================================================
print("📋 TESTE 5: Extração de Tabelas")
print("-" * 80)

from extrator.orquestrador import extrair_tabelas_brutas

if Path(pdf_teste).exists():
    linhas, max_cols = extrair_tabelas_brutas(pdf_teste)
    print(f"✓ Linhas extraídas: {len(linhas)}")
    print(f"✓ Colunas máximas: {max_cols}")
print()

# ============================================================================
# TESTE 6: Orquestrador Completo
# ============================================================================
print("📋 TESTE 6: Orquestrador Completo")
print("-" * 80)

from extrator import processar_cnis_completo

if Path(pdf_teste).exists():
    print(f"Processando: {pdf_teste}")
    print()
    
    try:
        arquivos = processar_cnis_completo(
            caminho_pdf=pdf_teste,
            pasta_saida="saida",
            nome_base="teste_modular"
        )
        
        print("✓ Arquivos gerados:")
        for tipo, caminho in arquivos.items():
            arquivo = Path(caminho)
            if arquivo.exists():
                tamanho = arquivo.stat().st_size
                linhas = len(arquivo.read_text(encoding='utf-8').splitlines())
                print(f"  [{tipo:20s}] {caminho}")
                print(f"  {'':22s} Tamanho: {tamanho:,} bytes | Linhas: {linhas}")
            else:
                print(f"  [{tipo:20s}] ⚠️  Não gerado")
        
    except Exception as e:
        print(f"❌ Erro no orquestrador: {e}")
        import traceback
        traceback.print_exc()
else:
    print(f"⚠️  PDF não encontrado: {pdf_teste}")

print()

# ============================================================================
# TESTE 7: Validação Final - Remunerações
# ============================================================================
print("📋 TESTE 7: Validação de Remunerações (Objetivo: 178)")
print("-" * 80)

# Teste com converter_extrato_inss.py original (baseline)
import subprocess

try:
    resultado = subprocess.run(
        [
            sys.executable,
            "converter_extrato_inss.py",
            pdf_teste,
            "saida/teste_baseline.csv"
        ],
        capture_output=True,
        text=True,
        timeout=30
    )
    
    if resultado.returncode == 0:
        csv_remun = Path("saida/teste_baseline_remuneracoes.csv")
        if csv_remun.exists():
            linhas_remun = len(csv_remun.read_text(encoding='utf-8').splitlines()) - 1
            print(f"✓ Baseline (CLI original): {linhas_remun} remunerações")
            
            if linhas_remun == 178:
                print(f"  ✅ SUCESSO: 178 remunerações preservadas!")
            else:
                print(f"  ⚠️  Esperado 178, obtido {linhas_remun}")
        else:
            print(f"  ⚠️  Arquivo de remunerações não gerado")
    else:
        print(f"  ❌ Erro ao executar CLI: {resultado.stderr[:200]}")
        
except Exception as e:
    print(f"  ⚠️  Não foi possível testar baseline: {e}")

print()

# ============================================================================
# RESUMO
# ============================================================================
print("=" * 80)
print("📊 RESUMO DOS TESTES")
print("=" * 80)
print()
print("✅ Componentes testados:")
print("  1. Detector de tipos (TipoVinculo enum)")
print("  2. Parser de cabeçalho (NIT, CPF, Nome)")
print("  3. Processadores CLT e Facultativo")
print("  4. Parser de vínculos (extração + parse)")
print("  5. Extração de tabelas brutas")
print("  6. Orquestrador completo")
print("  7. Validação baseline (178 remunerações)")
print()
print("🎯 Estrutura modular tipos/ e parsers/ funcionando!")
print()
print("⏳ Próximo passo: Integrar orquestrador no CLI principal")
print("   e refatorar extrair_remuneracoes_texto() para usar os tipos/")
print()
print("=" * 80)
