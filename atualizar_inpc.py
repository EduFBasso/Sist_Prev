#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para atualizar índices INPC do Banco Central do Brasil
Série 188 - INPC (Índice Nacional de Preços ao Consumidor)
Gera arquivo CSV com fatores de correção monetária para o sistema ERP_Prev
"""

import json
import os
import sys
from pathlib import Path
from datetime import datetime
from urllib.request import urlopen, Request
from urllib.error import URLError, HTTPError

def baixar_dados_bcb():
    """Baixa dados da série 188 (INPC) do BCB"""
    url = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.188/dados?formato=json"
    
    try:
        print("Conectando à API do Banco Central...")
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

def calcular_fatores(dados, mes_base=None):
    """
    Calcula fatores de correção monetária
    mes_base: competência base (None = usar último mês disponível)
    Retorna dicionário {competencia: fator}
    """
    # Converter dados para dicionário {competencia: valor_inpc}
    # Usar apenas o ULTIMO registro de cada mês (fim do mês)
    # Ignorar valores inválidos (zero, negativos, None)
    
    # Primeiro, organizar por mês/ano
    dados_por_mes = {}
    registros_invalidos = 0
    
    for registro in dados:
        comp_api = registro['data']  # formato DD/MM/YYYY da API
        partes = comp_api.split('/')
        dia = int(partes[0])
        mes = partes[1]
        ano = partes[2]
        comp = f"{mes}/{ano}"  # MM/YYYY
        
        try:
            valor = float(registro['valor'])
            # Filtrar valores muito baixos (< 0.5) que são erros de dados do BCB
            # e valores negativos ou zero
            if valor >= 0.5:  # Limite mínimo razoável para INPC mensal
                # Guardar o registro com maior dia (fim do mês)
                if comp not in dados_por_mes or dia > dados_por_mes[comp]['dia']:
                    dados_por_mes[comp] = {'dia': dia, 'valor': valor}
            else:
                registros_invalidos += 1
        except (ValueError, TypeError):
            registros_invalidos += 1
    
    # Criar dicionário final com valores do fim de cada mês
    inpc_dict = {comp: info['valor'] for comp, info in dados_por_mes.items()}
    
    if registros_invalidos > 0:
        print(f"⚠️  {registros_invalidos} registros com valores inválidos foram ignorados")
    
    # Se não especificou mês base, usar o último disponível
    if mes_base is None:
        # Ordenar competências por ano/mês
        competencias = sorted(inpc_dict.keys(), 
                            key=lambda x: (x.split('/')[1], x.split('/')[0]))
        mes_base = competencias[-1]  # Última competência
        print(f"\n✓ Usando último mês disponível como base: {mes_base}")
    
    print(f"\nCalculando fatores com base em {mes_base}...")
    
    # Verificar se mês base existe
    if mes_base not in inpc_dict:
        print(f"✗ Mês base {mes_base} não encontrado nos dados!")
        return None
    
    # IMPORTANTE: Os valores da API são VARIAÇÕES PERCENTUAIS MENSAIS
    # Precisamos ACUMULAR para criar índices comparáveis
    # Índice acumulado = produto de (1 + variação/100) ao longo do tempo
    
    print("Calculando índices acumulados...")
    indices_acumulados = {}
    indice_atual = 100.0  # Começar com base 100 em maio/1979
    
    # Ordenar competências cronologicamente
    competencias_ordenadas = sorted(inpc_dict.keys(), 
                                   key=lambda x: (x.split('/')[1], x.split('/')[0]))
    
    for comp in competencias_ordenadas:
        variacao_percentual = inpc_dict[comp]
        # Acumular: Índice novo = Índice anterior × (1 + variação/100)
        indice_atual = indice_atual * (1 + variacao_percentual / 100)
        indices_acumulados[comp] = indice_atual
    
    # Usar o índice acumulado do mês base
    indice_base = indices_acumulados[mes_base]
    
    if indice_base <= 0:
        print(f"✗ Índice acumulado do mês base é inválido: {indice_base}")
        return None
    
    print(f"Índice acumulado base ({mes_base}): {indice_base:.2f}")
    
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
    exemplos = [('01/1995', 'Janeiro/1995'), ('12/1999', 'Dezembro/1999'), 
                ('12/2020', 'Dezembro/2020'), (mes_base, f'{mes_base} (BASE)')]
    print("\n--- Exemplos de fatores (devem ser > 1 para datas antigas) ---")
    for comp, desc in exemplos:
        if comp in fatores:
            print(f"{desc}: {fatores[comp]:.6f}x")
    
    return fatores

def _pasta_base() -> Path:
    """Diretório base do programa (funciona em .py e .exe empacotado)."""
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        # Padrão do projeto: executáveis ficam em "bin/" dentro da pasta do Excel.
        # Nesse caso, a pasta base correta é o PAI de "bin".
        if exe_dir.name.lower() == "bin":
            return exe_dir.parent
        return exe_dir
    return Path(__file__).resolve().parent


def salvar_csv(fatores, arquivo: str | None = None):
    """Salva fatores em arquivo CSV"""
    try:
        # Sempre gravar dentro da pasta do programa (evita depender do "cwd")
        if not arquivo:
            arquivo_path = _pasta_base() / "saida" / "inpc_fatores.csv"
        else:
            arquivo_path = Path(arquivo)

        # Criar pasta saida se não existir
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
        
        # Mostrar alguns exemplos
        print("\n--- Exemplos de fatores calculados ---")
        exemplos = [
            ("01/1995", "Janeiro/1995"),
            ("12/1999", "Dezembro/1999"),
            ("06/2010", "Junho/2010"),
            ("12/2020", "Dezembro/2020"),
            ("12/2025", "Dezembro/2025 (BASE)")
        ]
        
        for comp, descricao in exemplos:
            if comp in fatores:
                print(f"{descricao}: {fatores[comp]:.6f}x")
        
        return True
        
    except Exception as e:
        print(f"✗ Erro ao salvar arquivo: {str(e)}")
        return False

def main():
    """Função principal"""
    print("=" * 60)
    print("ATUALIZAÇÃO DE ÍNDICES INPC - ERP_PREV")
    print("=" * 60)
    print(f"Data/Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
    
    # 1. Baixar dados
    dados = baixar_dados_bcb()
    if not dados:
        print("\n✗ FALHA: Não foi possível obter dados do BCB")
        input("\nPressione ENTER para sair...")
        return False
    
    # 2. Calcular fatores (None = usar último mês disponível)
    fatores = calcular_fatores(dados, mes_base=None)
    if not fatores:
        print("\n✗ FALHA: Erro ao calcular fatores")
        input("\nPressione ENTER para sair...")
        return False
    
    # 3. Salvar CSV
    sucesso = salvar_csv(fatores)
    
    print("\n" + "=" * 60)
    if sucesso:
        print("✓ ATUALIZAÇÃO CONCLUÍDA COM SUCESSO!")
    else:
        print("✗ ATUALIZAÇÃO FALHOU")
    print("=" * 60)
    
    input("\nPressione ENTER para sair...")
    return sucesso

if __name__ == "__main__":
    main()
