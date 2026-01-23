#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DETECTOR DE FORMATO CNIS - INSS

Analisa o PDF do CNIS e identifica qual formato está sendo usado.
Cada formato tem seu próprio extrator especializado.

Formatos conhecidos:
- FORMATO_LOGO: Logo INSS + título padrão (INSS / CNIS - Cadastro Nacional... / Extrato Previdenciário)
- FORMATO_SIMPLES: Sem logo ou variações no título (a implementar)

Autor: Sistema Previdenciário
Data: Janeiro 2026
"""

import sys
import pdfplumber
from pathlib import Path
import subprocess


def detectar_formato_cnis(pdf_path):
    """
    Analisa o PDF e detecta qual formato de CNIS está sendo usado.
    
    Args:
        pdf_path: Caminho para o arquivo PDF
        
    Returns:
        str: Nome do formato detectado ('FORMATO_LOGO', 'FORMATO_SIMPLES', etc.)
             ou None se não conseguir identificar
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # Analisar primeira página
            primeira_pagina = pdf.pages[0]
            texto = primeira_pagina.extract_text() or ""
            linhas = texto.split('\n')
            
            # Limpar linhas vazias
            linhas = [l.strip() for l in linhas if l.strip()]
            
            print("\n" + "="*70)
            print("🔍 DETECTANDO FORMATO DO CNIS")
            print("="*70)
            print(f"\n📄 Arquivo: {Path(pdf_path).name}")
            print(f"📑 Total de páginas: {len(pdf.pages)}")
            
            print("\n📋 Primeiras 10 linhas do documento:")
            for i, linha in enumerate(linhas[:10], 1):
                print(f"   {i:2d}. {linha[:80]}")
            
            # FORMATO_LOGO: Verificar padrão específico
            # - Deve ter "INSS" nas primeiras linhas
            # - Deve ter "CNIS - Cadastro Nacional de Informações Sociais"
            # - Deve ter "Extrato Previdenciário"
            
            texto_primeiras_linhas = '\n'.join(linhas[:15]).upper()
            
            tem_inss = "INSS" in texto_primeiras_linhas
            tem_cnis_completo = "CNIS" in texto_primeiras_linhas and "CADASTRO NACIONAL" in texto_primeiras_linhas
            tem_extrato_previdenciario = "EXTRATO PREVIDENCIÁRIO" in texto_primeiras_linhas or "EXTRATO PREVIDENCIARIO" in texto_primeiras_linhas
            
            print("\n🔎 Verificando assinaturas do formato:")
            print(f"   {'✓' if tem_inss else '✗'} INSS")
            print(f"   {'✓' if tem_cnis_completo else '✗'} CNIS - Cadastro Nacional de Informações Sociais")
            print(f"   {'✓' if tem_extrato_previdenciario else '✗'} Extrato Previdenciário")
            
            if tem_inss and tem_cnis_completo and tem_extrato_previdenciario:
                print("\n✅ Formato detectado: FORMATO_LOGO")
                print("   (Logo INSS + Título padrão + Estrutura com zonas)")
                return "FORMATO_LOGO"
            
            # Verificar se tem "Relações Previdenciárias" (comum em ambos formatos)
            if "RELAÇÕES PREVIDENCIÁRIAS" in texto_primeiras_linhas or "RELACOES PREVIDENCIARIAS" in texto_primeiras_linhas:
                print("\n⚠️  Formato detectado: FORMATO_DESCONHECIDO")
                print("   (Possui 'Relações Previdenciárias' mas não corresponde ao FORMATO_LOGO)")
                print("   💡 Este formato ainda não tem extrator específico implementado")
                return "FORMATO_DESCONHECIDO"
            
            print("\n❌ Formato NÃO IDENTIFICADO")
            print("   Este PDF não corresponde a nenhum formato conhecido")
            return None
            
    except Exception as e:
        print(f"\n❌ ERRO ao analisar PDF: {e}")
        return None


def executar_extrator(formato, pdf_path):
    """
    Executa o extrator apropriado para o formato detectado.
    
    Args:
        formato: Nome do formato ('FORMATO_LOGO', etc.)
        pdf_path: Caminho para o arquivo PDF
        
    Returns:
        bool: True se extração foi bem sucedida, False caso contrário
    """
    print("\n" + "="*70)
    print("🚀 EXECUTANDO EXTRATOR ESPECÍFICO")
    print("="*70)
    
    if formato == "FORMATO_LOGO":
        print(f"\n▶️  Chamando: converter_extrato_inss.py")
        print(f"📄 Arquivo: {pdf_path}")
        print()
        
        try:
            # Executar o extrator específico
            resultado = subprocess.run(
                [sys.executable, "converter_extrato_inss.py", pdf_path],
                check=True,
                capture_output=False,
                text=True
            )
            return resultado.returncode == 0
        except subprocess.CalledProcessError as e:
            print(f"\n❌ ERRO na extração: {e}")
            return False
        except FileNotFoundError:
            print(f"\n❌ ERRO: Arquivo converter_extrato_inss.py não encontrado")
            return False
    
    elif formato == "FORMATO_DESCONHECIDO":
        print("\n⚠️  Este formato ainda não possui extrator implementado")
        print("💡 Sugestão: Crie um novo extrator especializado para este formato")
        print("   Exemplo: converter_extrato_inss_formato_simples.py")
        return False
    
    else:
        print(f"\n❌ Formato '{formato}' não possui extrator implementado")
        return False


def main():
    """Função principal"""
    print("\n" + "="*70)
    print("🎯 DETECTOR E ORQUESTRADOR DE EXTRATORES CNIS")
    print("="*70)
    print("\nArquitetura:")
    print("  1. Detecta formato do PDF (LOGO, SIMPLES, etc.)")
    print("  2. Chama extrator específico para aquele formato")
    print("  3. Evita misturar lógicas de formatos diferentes")
    print()
    
    if len(sys.argv) < 2:
        print("❌ Uso: python detector_formato_cnis.py <caminho_do_pdf>")
        print()
        print("Exemplo:")
        print("  python detector_formato_cnis.py cnis/CNIS_JOAO_CARLOS.pdf")
        sys.exit(1)
    
    pdf_path = sys.argv[1]
    
    # Verificar se arquivo existe
    if not Path(pdf_path).exists():
        print(f"❌ ERRO: Arquivo não encontrado: {pdf_path}")
        sys.exit(1)
    
    # Detectar formato
    formato = detectar_formato_cnis(pdf_path)
    
    if formato is None:
        print("\n❌ Não foi possível identificar o formato do PDF")
        print("💡 Verifique se o arquivo é um extrato CNIS válido")
        sys.exit(1)
    
    # Executar extrator apropriado
    sucesso = executar_extrator(formato, pdf_path)
    
    if sucesso:
        print("\n" + "="*70)
        print("✅ EXTRAÇÃO CONCLUÍDA COM SUCESSO!")
        print("="*70)
        sys.exit(0)
    else:
        print("\n" + "="*70)
        print("❌ FALHA NA EXTRAÇÃO")
        print("="*70)
        sys.exit(1)


if __name__ == "__main__":
    main()
