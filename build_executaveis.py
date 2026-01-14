#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de Build - Gera executáveis .exe para distribuição ao cliente
Usa PyInstaller para converter scripts Python em executáveis standalone

Executar APENAS NO AMBIENTE DE DESENVOLVIMENTO (com Python instalado)
"""

import os
import sys
import shutil
import subprocess
from datetime import datetime

# Configurações
SCRIPTS_PARA_COMPILAR = [
    "atualizar_inpc.py",
    "converter_extrato_inss.py"
]

PASTA_BIN = "bin"
PASTA_DIST = "dist"
PASTA_BUILD = "build"
VERSAO = "1.0.0"

def print_header(mensagem):
    """Imprime cabeçalho formatado"""
    print("\n" + "=" * 70)
    print(f"  {mensagem}")
    print("=" * 70)

def verificar_pyinstaller():
    """Verifica se PyInstaller está instalado"""
    try:
        resultado = subprocess.run(
            ["pyinstaller", "--version"],
            capture_output=True,
            text=True,
            check=True
        )
        print(f"✓ PyInstaller encontrado: {resultado.stdout.strip()}")
        return True
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("✗ PyInstaller não encontrado!")
        print("\nInstale com: pip install pyinstaller")
        return False

def limpar_diretorios():
    """Remove diretórios de build anteriores"""
    print("\nLimpando diretórios antigos...")
    
    for pasta in [PASTA_BUILD, PASTA_DIST]:
        if os.path.exists(pasta):
            shutil.rmtree(pasta)
            print(f"  ✓ Removido: {pasta}/")
    
    # Limpar .spec files
    for arquivo in os.listdir("."):
        if arquivo.endswith(".spec"):
            os.remove(arquivo)
            print(f"  ✓ Removido: {arquivo}")

def compilar_script(script_name):
    """Compila um script Python para .exe"""
    if not os.path.exists(script_name):
        print(f"✗ Script não encontrado: {script_name}")
        return False
    
    print(f"\n  Compilando: {script_name}")
    print(f"  Modo: --onefile --console")
    
    try:
        # Comando PyInstaller
        comando = [
            "pyinstaller",
            "--onefile",        # Gerar um único .exe
            "--console",        # Manter janela de console (debug)
            "--clean",          # Limpar cache
            script_name
        ]
        
        resultado = subprocess.run(
            comando,
            capture_output=True,
            text=True,
            check=True
        )
        
        print(f"  ✓ Compilado com sucesso!")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"  ✗ Erro na compilação:")
        print(f"     {e.stderr}")
        return False

def mover_executaveis():
    """Move executáveis da pasta dist/ para bin/"""
    if not os.path.exists(PASTA_DIST):
        print(f"✗ Pasta {PASTA_DIST}/ não encontrada")
        return False
    
    # Criar pasta bin/ se não existir
    if not os.path.exists(PASTA_BIN):
        os.makedirs(PASTA_BIN)
    
    print(f"\nMovendo executáveis para {PASTA_BIN}/...")
    
    sucesso = True
    for arquivo in os.listdir(PASTA_DIST):
        if arquivo.endswith(".exe"):
            origem = os.path.join(PASTA_DIST, arquivo)
            destino = os.path.join(PASTA_BIN, arquivo)
            
            # Sobrescrever se existir
            if os.path.exists(destino):
                os.remove(destino)
            
            shutil.move(origem, destino)
            
            # Verificar tamanho
            tamanho = os.path.getsize(destino) / (1024 * 1024)  # MB
            print(f"  ✓ {arquivo} ({tamanho:.1f} MB)")
    
    return sucesso

def criar_arquivo_versao():
    """Cria arquivo com informações de versão"""
    versao_path = os.path.join(PASTA_BIN, "VERSAO.txt")
    
    with open(versao_path, 'w', encoding='utf-8') as f:
        f.write(f"ERP_PREV - Sistema de Previdência\n")
        f.write(f"Versão: {VERSAO}\n")
        f.write(f"Build: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
        f.write(f"\nExecutáveis incluídos:\n")
        
        for arquivo in sorted(os.listdir(PASTA_BIN)):
            if arquivo.endswith(".exe"):
                tamanho = os.path.getsize(os.path.join(PASTA_BIN, arquivo)) / (1024 * 1024)
                f.write(f"  - {arquivo} ({tamanho:.1f} MB)\n")
    
    print(f"\n✓ Arquivo de versão criado: {versao_path}")

def validar_build():
    """Valida se todos os executáveis foram gerados"""
    print("\nValidando build...")
    
    todos_ok = True
    for script in SCRIPTS_PARA_COMPILAR:
        exe_name = script.replace(".py", ".exe")
        exe_path = os.path.join(PASTA_BIN, exe_name)
        
        if os.path.exists(exe_path):
            print(f"  ✓ {exe_name}")
        else:
            print(f"  ✗ {exe_name} NÃO ENCONTRADO!")
            todos_ok = False
    
    return todos_ok

def exibir_resumo():
    """Exibe resumo final do build"""
    print_header("BUILD CONCLUÍDO")
    
    print("\n📦 ESTRUTURA PARA DISTRIBUIÇÃO:")
    print(f"\n  📁 {os.path.basename(os.getcwd())}/")
    print(f"    ├─ erp_prev.xlsm")
    print(f"    ├─ 📁 bin/")
    
    for arquivo in sorted(os.listdir(PASTA_BIN)):
        if arquivo.endswith(".exe"):
            tamanho = os.path.getsize(os.path.join(PASTA_BIN, arquivo)) / (1024 * 1024)
            print(f"    │   ├─ {arquivo} ({tamanho:.1f} MB)")
    
    print(f"    │   └─ VERSAO.txt")
    print(f"    └─ 📁 saida/")
    
    # Tamanho total
    tamanho_total = sum(
        os.path.getsize(os.path.join(PASTA_BIN, f))
        for f in os.listdir(PASTA_BIN)
        if f.endswith(".exe")
    ) / (1024 * 1024)
    
    print(f"\n  Tamanho total dos executáveis: {tamanho_total:.1f} MB")
    
    print("\n📧 PRÓXIMOS PASSOS:")
    print("  1. Testar executáveis no Windows (Parallels)")
    print("  2. Verificar CHECKLIST_ENVIO.md")
    print("  3. Comprimir pasta para envio por e-mail")
    print("  4. Enviar com INSTALACAO_CLIENTE.md")

def main():
    """Função principal"""
    print_header(f"BUILD - ERP_PREV v{VERSAO}")
    print(f"Data/Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    
    # 1. Verificar PyInstaller
    if not verificar_pyinstaller():
        return 1
    
    # 2. Limpar diretórios antigos
    limpar_diretorios()
    
    # 3. Compilar cada script
    print_header("COMPILANDO SCRIPTS")
    
    for script in SCRIPTS_PARA_COMPILAR:
        if not compilar_script(script):
            print(f"\n✗ Falha ao compilar {script}")
            return 1
    
    # 4. Mover executáveis
    print_header("ORGANIZANDO EXECUTÁVEIS")
    if not mover_executaveis():
        print("✗ Falha ao mover executáveis")
        return 1
    
    # 5. Criar arquivo de versão
    criar_arquivo_versao()
    
    # 6. Validar build
    if not validar_build():
        print("\n✗ Build inválido - alguns executáveis faltando")
        return 1
    
    # 7. Limpar arquivos temporários
    print("\nLimpando arquivos temporários...")
    limpar_diretorios()
    
    # 8. Resumo final
    exibir_resumo()
    
    print("\n✓ BUILD CONCLUÍDO COM SUCESSO!\n")
    return 0

if __name__ == "__main__":
    sys.exit(main())
