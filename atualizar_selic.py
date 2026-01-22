#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para atualizar índices SELIC do Banco Central do Brasil
Série 1178 - SELIC acumulada mensal (%)
Gera arquivo CSV com fatores de correção monetária para o sistema ERP_Prev

IMPORTANTE: A série 1178 é SELIC ACUMULADA no mês (não taxa diária)
"""

import json
import os
import sys
from pathlib import Path
from datetime import datetime, timedelta
from urllib.request import urlopen, Request
from urllib.error import URLError, HTTPError

def baixar_dados_bcb_janela(data_inicial, data_final):
    """
    Baixa dados da série 1178 (SELIC acumulada mensal) do BCB
    data_inicial e data_final em formato DD/MM/AAAA
    """
    url = (f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.1178/dados?"
           f"formato=json&dataInicial={data_inicial}&dataFinal={data_final}")
    
    try:
        print(f"Buscando {data_inicial} até {data_final}...", end=" ")
        req = Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        
        with urlopen(req, timeout=30) as response:
            dados = json.loads(response.read().decode('utf-8'))
        
        print(f"✓ {len(dados)} registros")
        return dados
    
    except HTTPError as e:
        print(f"✗ Erro HTTP {e.code}: {e.reason}")
        return []
    except URLError as e:
        print(f"✗ Erro de conexão: {e.reason}")
        return []
    except Exception as e:
        print(f"✗ Erro: {str(e)}")
        return []

def baixar_todos_dados_selic():
    """
    Baixa dados históricos completos da SELIC
    Como a API limita a 10 anos por janela, faz múltiplas requisições
    """
    print("=" * 70)
    print("ATUALIZANDO SELIC - Série 1178 (SELIC acumulada mensal)")
    print("=" * 70)
    print("\nConectando à API do Banco Central...")
    print("(A API limita consultas a 10 anos por vez, fazendo múltiplas requisições)\n")
    
    # Buscar desde 1986 (início da série 1178) até hoje
    # Dividir em janelas de 10 anos
    
    ano_inicial = 1986
    ano_final = datetime.now().year
    
    todos_dados = []
    
    ano_atual = ano_inicial
    while ano_atual <= ano_final:
        # Define janela de 10 anos (ou menos se for a última)
        data_inicio = f"01/01/{ano_atual}"
        ano_fim = min(ano_atual + 9, ano_final)
        
        # Se for o ano atual, usar data de hoje
        if ano_fim == ano_final:
            data_fim = datetime.now().strftime("%d/%m/%Y")
        else:
            data_fim = f"31/12/{ano_fim}"
        
        # Buscar dados da janela
        dados_janela = baixar_dados_bcb_janela(data_inicio, data_fim)
        todos_dados.extend(dados_janela)
        
        # Próxima janela
        ano_atual = ano_fim + 1
    
    print(f"\n✓ Total de {len(todos_dados)} registros obtidos")
    return todos_dados

def processar_selic_mensal(dados):
    """
    Processa dados da SELIC para obter valores mensais
    A série 1178 já é ACUMULADA no mês, então pegamos o ÚLTIMO valor de cada mês
    
    IMPORTANTE: Os valores da API já são % acumulado do mês inteiro
    Exemplo: Janeiro/1987 = 1100 significa 1100% acumulado no mês
    """
    # Organizar por mês/ano
    dados_por_mes = {}
    registros_invalidos = 0
    
    print("\nProcessando dados mensais...")
    
    for registro in dados:
        comp_api = registro['data']  # formato DD/MM/YYYY da API
        partes = comp_api.split('/')
        dia = int(partes[0])
        mes = partes[1]
        ano = partes[2]
        comp = f"{mes}/{ano}"  # MM/YYYY
        
        try:
            # A API retorna valores já como percentual (%)
            # Janeiro/1987 = 1100 significa 1100% no mês
            valor = float(registro['valor'])
            
            # Valores válidos: podem ser muito altos nos anos 80/90 (hiperinflação)
            # Filtrar apenas zeros ou negativos (erros)
            if valor >= 0:
                # Guardar o registro com maior dia (fim do mês)
                if comp not in dados_por_mes or dia > dados_por_mes[comp]['dia']:
                    dados_por_mes[comp] = {'dia': dia, 'valor': valor}
            else:
                registros_invalidos += 1
        except (ValueError, TypeError):
            registros_invalidos += 1
    
    # Criar dicionário final com valores do fim de cada mês
    selic_dict = {comp: info['valor'] for comp, info in dados_por_mes.items()}
    
    if registros_invalidos > 0:
        print(f"⚠️  {registros_invalidos} registros com valores inválidos foram ignorados")
    
    print(f"✓ {len(selic_dict)} competências processadas")
    
    # Mostrar alguns exemplos para contexto histórico
    print("\n--- CONTEXTO HISTÓRICO (valores são CORRETOS!) ---")
    exemplos = [
        ('01/1987', 'Janeiro/1987 (Hiperinflação)'),
        ('12/1994', 'Dezembro/1994 (Pré-Plano Real)'),
        ('01/2000', 'Janeiro/2000 (Pós-Real)'),
        ('12/2020', 'Dezembro/2020'),
    ]
    
    for comp, desc in exemplos:
        if comp in selic_dict:
            valor = selic_dict[comp]
            print(f"{desc}: {valor:.2f}% ao mês")
            if valor > 100:
                print(f"  → Isso equivale a {valor:.0f}% acumulado no mês (hiperinflação!)")
    
    return selic_dict

def calcular_fatores_selic(selic_dict, mes_base=None):
    """
    Calcula fatores de correção monetária baseados na SELIC
    
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
    
    # IMPORTANTE: Os valores da API são VARIAÇÕES PERCENTUAIS MENSAIS ACUMULADAS
    # Precisamos ACUMULAR para criar índices comparáveis
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
        ('01/1995', 'Janeiro/1995 (Pré-Real)'), 
        ('12/1999', 'Dezembro/1999'), 
        ('12/2010', 'Dezembro/2010'),
        ('12/2020', 'Dezembro/2020'), 
        (mes_base, f'{mes_base} (BASE)')
    ]
    print("\n--- Exemplos de fatores (valores antigos têm fator >> 1) ---")
    for comp, desc in exemplos:
        if comp in fatores:
            fator = fatores[comp]
            print(f"{desc}: {fator:,.2f}x")
            # Mostrar quanto R$ 1000 de 'comp' vale hoje
            if comp != mes_base:
                valor_corrigido = 1000 * fator
                print(f"  → R$ 1.000 de {comp} = R$ {valor_corrigido:,.2f} em {mes_base}")
    
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
            arquivo_path = _pasta_base() / "saida" / "selic_fatores.csv"
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
        print(f"✓ Caminho: {arquivo_path}")
        
        return True
    
    except Exception as e:
        print(f"✗ Erro ao salvar arquivo: {str(e)}")
        return False

def main():
    """Função principal"""
    try:
        # Baixar dados
        dados = baixar_todos_dados_selic()
        
        if not dados:
            print("\n✗ Nenhum dado foi obtido. Verifique sua conexão com a internet.")
            input("\nPressione ENTER para sair...")
            return 1
        
        # Processar dados mensais
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
        
        # Salvar CSV
        if salvar_csv(fatores):
            print("\n" + "=" * 70)
            print("✓ ATUALIZAÇÃO CONCLUÍDA COM SUCESSO!")
            print("=" * 70)
            print("\nImporte o arquivo 'selic_fatores.csv' no Excel para usar os fatores.")
            print("\nNOTA: Valores dos anos 80/90 são REALMENTE altos devido à hiperinflação.")
            print("      Exemplo: Janeiro/1987 = 1100% ao mês era a realidade econômica.")
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
