#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para atualizar índices SELIC - Série 4390
Série 4390 - SELIC acumulada mensalmente (% a.m.)
ESTA É A SÉRIE CORRETA para cálculos previdenciários!

Diferenças entre as séries:
- Série 1178: Taxa SELIC diária (valores muito altos, não adequada)
- Série 4390: Taxa SELIC acumulada no mês (% a.m.) - USAR ESTA!
"""

import json
import os
import sys
from pathlib import Path
from datetime import datetime
from urllib.request import urlopen, Request
from urllib.error import URLError, HTTPError

def baixar_dados_bcb():
    """Baixa dados da série 4390 (SELIC acumulada mensal %)"""
    url = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.4390/dados?formato=json"
    
    try:
        print("=" * 70)
        print("ATUALIZANDO SELIC - Série 4390 (Taxa acumulada mensal %)")
        print("=" * 70)
        print("\nConectando à API do Banco Central...")
        
        req = Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        
        with urlopen(req, timeout=30) as response:
            dados = json.loads(response.read().decode('utf-8'))
        
        print(f"✓ {len(dados)} registros obtidos com sucesso")
        return dados
    
    except HTTPError as e:
        print(f"✗ Erro HTTP {e.code}: {e.reason}")
        return None
    except URLError as e:
        print(f"✗ Erro de conexão: {e.reason}")
        return None
    except Exception as e:
        print(f"✗ Erro inesperado: {str(e)}")
        return None

def processar_selic_mensal(dados):
    """
    Processa dados da SELIC série 4390
    Esta série já é MENSAL (um valor por mês)
    Valores são % acumulado no mês (razoáveis, não astronômicos)
    """
    # Converter dados para dicionário {competencia: valor}
    selic_dict = {}
    registros_invalidos = 0
    
    print("\nProcessando dados mensais...")
    
    for registro in dados:
        comp_api = registro['data']  # formato DD/MM/YYYY da API
        partes = comp_api.split('/')
        mes = partes[1]
        ano = partes[2]
        comp = f"{mes}/{ano}"  # MM/YYYY
        
        try:
            # A API retorna valores como percentual mensal (%)
            valor = float(registro['valor'])
            
            # Valores válidos: SELIC mensal típica é 0.3% a 2% ao mês
            # Nos anos 80/90 pode chegar a 30-50% ao mês (hiperinflação)
            if valor >= 0:
                selic_dict[comp] = valor
            else:
                registros_invalidos += 1
        except (ValueError, TypeError):
            registros_invalidos += 1
    
    if registros_invalidos > 0:
        print(f"⚠️  {registros_invalidos} registros com valores inválidos foram ignorados")
    
    print(f"✓ {len(selic_dict)} competências processadas")
    
    # Mostrar alguns exemplos para contexto
    print("\n--- CONTEXTO HISTÓRICO (Série 4390 - valores razoáveis) ---")
    exemplos = [
        ('01/1995', 'Janeiro/1995 (Hiperinflação)'),
        ('07/1994', 'Julho/1994 (Mês do Plano Real)'),
        ('12/1999', 'Dezembro/1999'),
        ('12/2010', 'Dezembro/2010'),
        ('12/2020', 'Dezembro/2020'),
    ]
    
    for comp, desc in exemplos:
        if comp in selic_dict:
            valor = selic_dict[comp]
            print(f"{desc}: {valor:.4f}% ao mês")
    
    return selic_dict

def calcular_fatores_selic(selic_dict, mes_base=None):
    """
    Calcula fatores de correção monetária baseados na SELIC
    Série 4390 - acumulação mensal razoável
    
    Retorna dicionário {competencia: fator}
    """
    # Se não especificou mês base, usar o último disponível
    if mes_base is None:
        # Ordenar competências por ano/mês
        competencias = sorted(selic_dict.keys(), 
                            key=lambda x: (x.split('/')[1], x.split('/')[0]))
        mes_base = competencias[-1]  # Última competência
        print(f"\n✓ Usando último mês disponível como base: {mes_base}")
    
    print(f"\nCalculando fatores de correção com base em {mes_base}...")
    
    # Verificar se mês base existe
    if mes_base not in selic_dict:
        print(f"✗ Mês base {mes_base} não encontrado nos dados!")
        return None
    
    # Calcular índices acumulados
    # Índice acumulado = produto de (1 + variação/100) ao longo do tempo
    
    print("Calculando índices acumulados...")
    indices_acumulados = {}
    indice_atual = 100.0  # Começar com base 100
    
    # Ordenar competências cronologicamente
    competencias_ordenadas = sorted(selic_dict.keys(), 
                                   key=lambda x: (x.split('/')[1], x.split('/')[0]))
    
    for comp in competencias_ordenadas:
        variacao_percentual = selic_dict[comp]
        # Acumular: Índice novo = Índice anterior × (1 + variação/100)
        indice_atual = indice_atual * (1 + variacao_percentual / 100)
        indices_acumulados[comp] = indice_atual
    
    # Usar o índice acumulado do mês base
    indice_base = indices_acumulados[mes_base]
    
    if indice_base <= 0:
        print(f"✗ Índice acumulado do mês base é inválido: {indice_base}")
        return None
    
    print(f"Índice acumulado base ({mes_base}): {indice_base:,.2f}")
    
    # Calcular fatores de correção
    # Fator = Índice_atual / Índice_passado
    # Valores antigos têm fator > 1 (aumentam)
    # Valores recentes têm fator ≈ 1
    fatores = {}
    
    for comp in indices_acumulados:
        indice_mes = indices_acumulados[comp]
        if indice_mes > 0:
            # Fator para trazer valor de 'comp' para 'mes_base'
            fator = indice_base / indice_mes
            fatores[comp] = fator
    
    print(f"✓ Fatores calculados para {len(fatores)} competências")
    
    # Mostrar alguns exemplos para validação
    exemplos = [
        ('01/1995', 'Janeiro/1995'), 
        ('12/1999', 'Dezembro/1999'), 
        ('12/2010', 'Dezembro/2010'),
        ('12/2020', 'Dezembro/2020'), 
        (mes_base, f'{mes_base} (BASE)')
    ]
    print("\n--- Exemplos de fatores de correção ---")
    for comp, desc in exemplos:
        if comp in fatores:
            fator = fatores[comp]
            print(f"{desc}: {fator:,.2f}x")
            # Mostrar quanto R$ 1000 de 'comp' vale hoje
            if comp != mes_base and fator < 1000000:  # Só mostrar se razoável
                valor_corrigido = 1000 * fator
                print(f"  → R$ 1.000,00 de {comp} = R$ {valor_corrigido:,.2f} em {mes_base}")
    
    return fatores

def _pasta_base() -> Path:
    """Diretório base do programa"""
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        if exe_dir.name.lower() == "bin":
            return exe_dir.parent
        return exe_dir
    return Path(__file__).resolve().parent

def salvar_csv(fatores, arquivo: str | None = None):
    """Salva fatores em arquivo CSV"""
    try:
        if not arquivo:
            arquivo_path = _pasta_base() / "saida" / "selic_fatores.csv"
        else:
            arquivo_path = Path(arquivo)

        os.makedirs(os.path.dirname(str(arquivo_path)), exist_ok=True)
        
        print(f"\nGravando arquivo {arquivo_path}...")
        
        with open(arquivo_path, 'w', encoding='utf-8') as f:
            f.write("Competencia;Fator\n")
            
            # Ordenar por competência (MM/YYYY)
            competencias = sorted(fatores.keys(), 
                                key=lambda x: (x.split('/')[1], x.split('/')[0]))
            
            for comp in competencias:
                fator = fatores[comp]
                f.write(f"{comp};{fator:.6f}\n")
        
        print(f"✓ Arquivo salvo com {len(fatores)} registros")
        print(f"✓ Caminho: {arquivo_path}")
        
        return True
    
    except Exception as e:
        print(f"✗ Erro ao salvar arquivo: {str(e)}")
        return False

def salvar_para_excel(dados_brutos, arquivo: str | None = None):
    """
    Salva dados brutos no formato usado pela advogada
    Compatível com script VBA: Data | Valor | Fonte
    """
    try:
        if not arquivo:
            arquivo_path = _pasta_base() / "saida" / "selic_bruto_para_excel.csv"
        else:
            arquivo_path = Path(arquivo)

        os.makedirs(os.path.dirname(str(arquivo_path)), exist_ok=True)
        
        print(f"\nGravando arquivo para Excel (formato VBA): {arquivo_path}...")
        
        with open(arquivo_path, 'w', encoding='utf-8') as f:
            # Cabeçalho igual ao usado no VBA
            f.write("Data;Valor;Fonte\n")
            
            for registro in dados_brutos:
                data = registro['data']  # DD/MM/YYYY
                valor = float(registro['valor'])
                f.write(f"{data};{valor:.6f};BCB API\n")
        
        print(f"✓ Arquivo salvo com {len(dados_brutos)} registros")
        print(f"✓ Este arquivo pode ser importado diretamente no Excel")
        
        return True
    
    except Exception as e:
        print(f"✗ Erro ao salvar arquivo: {str(e)}")
        return False

def main():
    """Função principal"""
    try:
        # Baixar dados
        dados = baixar_dados_bcb()
        
        if not dados:
            print("\n✗ Nenhum dado foi obtido. Verifique sua conexão com a internet.")
            input("\nPressione ENTER para sair...")
            return 1
        
        # Salvar dados brutos (formato VBA)
        print("\n" + "=" * 70)
        print("SALVANDO DADOS BRUTOS (formato compatível com VBA)")
        print("=" * 70)
        salvar_para_excel(dados)
        
        # Processar dados mensais
        print("\n" + "=" * 70)
        print("CALCULANDO FATORES DE CORREÇÃO")
        print("=" * 70)
        selic_dict = processar_selic_mensal(dados)
        
        if not selic_dict:
            print("\n✗ Erro ao processar dados.")
            input("\nPressione ENTER para sair...")
            return 1
        
        # Calcular fatores (usar último mês disponível como base)
        fatores = calcular_fatores_selic(selic_dict)
        
        if not fatores:
            print("\n✗ Erro ao calcular fatores.")
            input("\nPressione ENTER para sair...")
            return 1
        
        # Salvar CSV com fatores
        if salvar_csv(fatores):
            print("\n" + "=" * 70)
            print("✓ ATUALIZAÇÃO CONCLUÍDA COM SUCESSO!")
            print("=" * 70)
            print("\nArquivos gerados:")
            print("  1. selic_fatores.csv - Fatores de correção calculados")
            print("  2. selic_bruto_para_excel.csv - Dados brutos (formato VBA)")
            print("\n💡 A série 4390 tem valores RAZOÁVEIS (não astronômicos)")
            print("   Janeiro/1995 ≈ 2-4% ao mês (não 1100%!)")
        else:
            print("\n✗ Falha ao salvar arquivo.")
            input("\nPressione ENTER para sair...")
            return 1
        
        input("\nPressione ENTER para sair...")
        return 0
    
    except KeyboardInterrupt:
        print("\n\n✗ Operação cancelada pelo usuário.")
        return 1
    except Exception as e:
        print(f"\n✗ Erro inesperado: {str(e)}")
        import traceback
        traceback.print_exc()
        input("\nPressione ENTER para sair...")
        return 1

if __name__ == "__main__":
    sys.exit(main())
