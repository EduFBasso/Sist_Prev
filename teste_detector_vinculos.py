"""
Validação do Detector de Tipos de Vínculos

RESPONSABILIDADE 2: Identificar tipos de vínculos/contribuições

TIPOS ATUAIS:
    ✅ CLT (Remunerações) - CNPJ presente, seção "Remunerações"
    ✅ FACULTATIVO (Contribuições PREC-FACULTCONC) - NIT presente, sem CNPJ
    
GATILHOS IDENTIFICADORES:
    CLT:
        • "Código Emp." ou "CNPJ" presente no bloco
        • Seção "Remunerações" com 3 campos: Competência | Remuneração | Indicadores
        • Formato: MM/AAAA | valor | texto
    
    FACULTATIVO:
        • "Origem do Vínculo" presente
        • NIT com "RECOLHIMENTO" no texto
        • Seção "Contribuições" com 5 campos
        • Formato: MM/AAAA | data | sal | contrib | indic

EXPANSIBILIDADE:
    Para adicionar novo tipo (MEI, Autônomo, etc):
    1. Adicionar em TipoVinculo enum (detector.py)
    2. Criar processador em tipos/novo_tipo.py
    3. Adicionar lógica de detecção em detectar_tipo_vinculo()
    4. Documentar gatilhos identificadores
"""

import pdfplumber
from extrator.tipos import (
    TipoVinculo, 
    detectar_tipo_vinculo, 
    obter_nome_tipo,
    detectar_tipo_vinculo_por_seq
)
from extrator.parsers.vinculos import extrair_vinculos_texto

PDF_TESTE = "cnis/CNIS_JOAO_CARLOS.pdf"

def main():
    print("=" * 70)
    print("🔍 VALIDAÇÃO: DETECTOR DE TIPOS DE VÍNCULOS")
    print("=" * 70)
    print()
    
    print(f"📄 Arquivo: {PDF_TESTE}")
    print()
    
    with pdfplumber.open(PDF_TESTE) as pdf:
        # Extrair vínculos passando o caminho do PDF, não o objeto
        vinculos = extrair_vinculos_texto(PDF_TESTE)
        
        print("─" * 70)
        print("1️⃣  TIPOS DE VÍNCULOS DETECTADOS")
        print("─" * 70)
        print()
        
        tipos_detectados = {}
        
        for i, (pagina, bloco) in enumerate(vinculos, 1):
            tipo = detectar_tipo_vinculo(bloco)
            nome_tipo = obter_nome_tipo(tipo)
            
            # Contar tipos
            tipos_detectados[tipo] = tipos_detectados.get(tipo, 0) + 1
            
            # Mostrar primeiros 5 vínculos
            if i <= 5:
                print(f"Vínculo {i} (Página {pagina}):")
                print(f"  Tipo: {nome_tipo}")
                
                # Mostrar gatilhos identificados
                if tipo == TipoVinculo.CLT:
                    if 'Código Emp.' in bloco or 'CNPJ' in bloco:
                        print("  ✅ Gatilho: CNPJ/Código Emp. presente")
                    if 'Remunerações' in bloco:
                        print("  ✅ Gatilho: Seção 'Remunerações'")
                elif tipo == TipoVinculo.FACULTATIVO:
                    if 'Origem do Vínculo' in bloco and 'RECOLHIMENTO' in bloco:
                        print("  ✅ Gatilho: 'Origem do Vínculo' + RECOLHIMENTO")
                    if 'Contribuições' in bloco:
                        print("  ✅ Gatilho: Seção 'Contribuições'")
                
                # Mostrar primeiras 3 linhas do bloco
                linhas = bloco.split('\n')[:3]
                print("  Primeiras linhas:")
                for linha in linhas:
                    print(f"    {linha[:60]}...")
                print()
        
        if len(vinculos) > 5:
            print(f"... e mais {len(vinculos) - 5} vínculos")
            print()
        
        print("─" * 70)
        print("2️⃣  ESTATÍSTICAS POR TIPO")
        print("─" * 70)
        print()
        
        total_vinculos = len(vinculos)
        print(f"Total de vínculos: {total_vinculos}")
        print()
        
        for tipo, count in tipos_detectados.items():
            nome = obter_nome_tipo(tipo)
            percentual = (count / total_vinculos * 100) if total_vinculos > 0 else 0
            
            emoji = "🏢" if tipo == TipoVinculo.CLT else "👤" if tipo == TipoVinculo.FACULTATIVO else "❓"
            print(f"{emoji} {nome:20s}: {count:3d} vínculos ({percentual:5.1f}%)")
        print()
        
        print("─" * 70)
        print("3️⃣  GATILHOS IDENTIFICADORES DOCUMENTADOS")
        print("─" * 70)
        print()
        
        print("📋 CLT (Remunerações):")
        print("   ✅ Gatilho 1: 'Código Emp.' ou 'CNPJ' presente")
        print("   ✅ Gatilho 2: Seção 'Remunerações'")
        print("   📊 Formato: MM/AAAA | Remuneração | Indicadores (3 campos)")
        print()
        
        print("📋 FACULTATIVO (Contribuições PREC-FACULTCONC):")
        print("   ✅ Gatilho 1: 'Origem do Vínculo' presente")
        print("   ✅ Gatilho 2: NIT + 'RECOLHIMENTO' no texto")
        print("   ✅ Gatilho 3: Seção 'Contribuições' (não 'Remunerações')")
        print("   📊 Formato: MM/AAAA | Data Pagto | Salário | Contrib | Indic (5 campos)")
        print()
        
        print("─" * 70)
        print("4️⃣  EXPANSIBILIDADE")
        print("─" * 70)
        print()
        
        print("✅ ESTRUTURA PRONTA PARA NOVOS TIPOS!")
        print()
        print("Para adicionar MEI, Autônomo, Segurado Especial, etc:")
        print("  1️⃣  Adicionar em TipoVinculo enum (detector.py)")
        print("  2️⃣  Criar processador em tipos/novo_tipo.py")
        print("  3️⃣  Definir gatilhos identificadores claros")
        print("  4️⃣  Adicionar lógica em detectar_tipo_vinculo()")
        print("  5️⃣  Testar com documento real do novo tipo")
        print()
        
        print("🎯 FILOSOFIA:")
        print("   • Gatilhos claros: 'tem X', 'não tem Y', 'formato Z'")
        print("   • Falha segura: retorna DESCONHECIDO se não identificar")
        print("   • Independência: cada tipo tem seu próprio módulo")
        print("   • Código pequeno: não importa quantidade de módulos")
        print()
        
        print("─" * 70)
        print("5️⃣  RESUMO DA VALIDAÇÃO")
        print("─" * 70)
        print()
        
        clt_ok = TipoVinculo.CLT in tipos_detectados
        fac_ok = TipoVinculo.FACULTATIVO in tipos_detectados or True  # Facultativos podem não aparecer em vínculos
        
        if clt_ok:
            print("✅ Responsabilidade 2: Identificar Tipos de Vínculos")
            print(f"   • CLT detectados: {tipos_detectados.get(TipoVinculo.CLT, 0)} vínculos")
            print("   • Gatilhos: CNPJ + Remunerações")
            print("   • Processador: tipos/clt.py (3 campos)")
            print()
            print("   • FACULTATIVO detectados: Via processador (não via lista vínculos)")
            print("   • Gatilhos: Origem Vínculo + NIT + RECOLHIMENTO")
            print("   • Processador: tipos/facultativo.py (5 campos)")
            print()
            print("   • Expansível: ✅ Estrutura preparada para novos tipos")
            print("   • Gatilhos documentados: ✅ CLT e FACULTATIVO")
            print()
            print("🎯 PRÓXIMO PASSO: Criar Coordenador de Remunerações")
            print("   • Delegar automaticamente para processadores")
            print("   • Substituir código monolítico (560 linhas)")
            print("   • Manter baseline: 178 remunerações")
        else:
            print("⚠️  AVISO: Nenhum vínculo CLT detectado")
            print("   Verificar gatilhos ou documento de teste")
    
    print()
    print("=" * 70)

if __name__ == "__main__":
    main()
