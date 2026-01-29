#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Build para Advogado - Pacote completo do Sistema de Análise Previdenciária
Gera executável standalone e estrutura de distribuição
"""

import os
import sys
import shutil
import subprocess
from datetime import datetime
from pathlib import Path

# Configurações
VERSAO = "1.0.0"
PASTA_DIST = "C:\\dev\\dist_advogado"  # Caminho curto
PASTA_BUILD = "build"
NOME_EXECUTAVEL = "AnalisadorINSS"

# Arquivos/pastas necessários (apenas essenciais)
ESTRUTURA_DISTRIBUICAO = {
    'fat_inpc_selic': ['inpc_fatores.csv', 'selic_fatores.csv'],
    'info': [
        'GUIA_RAPIDO.md',
        'CALCULO_FATOR_DETALHADO.md'
    ],
    'saida': []  # Pasta vazia para output
}

# Arquivos de desenvolvimento para remover
ARQUIVOS_DESENVOLVIMENTO = [
    '__pycache__',
    '.venv',
    '.vscode',
    '*.pyc',
    '*.pyo',
    '*.spec',
    'debug_*.py',
    'teste_*.py',
    'CHANGELOG.txt',
    'README.txt'
]

def print_header(mensagem):
    """Cabeçalho formatado"""
    print("\n" + "=" * 80)
    print(f"  {mensagem}")
    print("=" * 80)

def verificar_dependencias():
    """Verifica se todas as dependências estão instaladas"""
    print("\n📦 Verificando dependências...")
    
    dependencias = {
        'PyInstaller': 'pyinstaller',
        'openpyxl': 'openpyxl',
        'customtkinter': 'customtkinter',
        'dateutil': 'python-dateutil'
    }
    
    faltando = []
    for modulo, pacote in dependencias.items():
        try:
            __import__(modulo)
            print(f"  ✓ {pacote}")
        except ImportError:
            print(f"  ✗ {pacote} - FALTANDO!")
            faltando.append(pacote)
    
    if faltando:
        print(f"\n❌ Instale as dependências faltando:")
        print(f"   pip install {' '.join([f.lower() for f in faltando])}")
        return False
    
    return True

def limpar_build_anterior():
    """Remove builds anteriores"""
    print("\n🧹 Limpando builds anteriores...")
    
    pastas_limpar = [PASTA_BUILD, PASTA_DIST, 'dist']
    for pasta in pastas_limpar:
        if os.path.exists(pasta):
            shutil.rmtree(pasta)
            print(f"  ✓ Removido: {pasta}/")
    
    # Limpar .spec files
    for arquivo in Path('.').glob('*.spec'):
        arquivo.unlink()
        print(f"  ✓ Removido: {arquivo}")

def criar_spec_file():
    """Cria arquivo .spec customizado para PyInstaller"""
    print("\n📝 Criando arquivo .spec...")
    
    spec_content = f'''# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['interface_cnis.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('fat_inpc_selic/inpc_fatores.csv', 'fat_inpc_selic'),
        ('fat_inpc_selic/selic_fatores.csv', 'fat_inpc_selic'),
    ],
    hiddenimports=[
        'openpyxl',
        'openpyxl.cell',
        'openpyxl.styles',
        'customtkinter',
        'dateutil',
        'dateutil.relativedelta',
        'tkinter',
    ],
    hookspath=[],
    hooksconfig={{}},
    runtime_hooks=[],
    excludes=['matplotlib', 'numpy', 'pandas', 'scipy'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='{NOME_EXECUTAVEL}',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Sem janela de console (GUI)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Adicionar ícone se existir
)
'''
    
    with open(f'{NOME_EXECUTAVEL}.spec', 'w', encoding='utf-8') as f:
        f.write(spec_content)
    
    print(f"  ✓ {NOME_EXECUTAVEL}.spec criado")
    return f'{NOME_EXECUTAVEL}.spec'

def compilar_executavel(spec_file):
    """Compila o executável usando PyInstaller"""
    print(f"\n🔨 Compilando {NOME_EXECUTAVEL}.exe...")
    print("  (Isso pode levar alguns minutos...)")
    
    try:
        comando = [
            sys.executable, '-m', 'PyInstaller',
            '--clean',
            '--noconfirm',
            spec_file
        ]
        
        resultado = subprocess.run(
            comando,
            capture_output=True,
            text=True,
            check=True
        )
        
        # Verificar se executável foi criado
        exe_path = os.path.join('dist', f'{NOME_EXECUTAVEL}.exe')
        if os.path.exists(exe_path):
            tamanho = os.path.getsize(exe_path) / (1024 * 1024)
            print(f"  ✓ Compilado com sucesso! ({tamanho:.1f} MB)")
            return exe_path
        else:
            print("  ✗ Executável não foi gerado")
            return None
            
    except subprocess.CalledProcessError as e:
        print(f"  ✗ Erro na compilação:")
        print(e.stderr)
        return None

def criar_estrutura_distribuicao(exe_path):
    """Cria estrutura de pastas para distribuição"""
    print(f"\n📁 Criando estrutura de distribuição em {PASTA_DIST}/...")
    
    # Criar pasta principal
    if os.path.exists(PASTA_DIST):
        shutil.rmtree(PASTA_DIST)
    os.makedirs(PASTA_DIST)
    
    # Copiar executável
    dest_exe = os.path.join(PASTA_DIST, f'{NOME_EXECUTAVEL}.exe')
    shutil.copy2(exe_path, dest_exe)
    tamanho = os.path.getsize(dest_exe) / (1024 * 1024)
    print(f"  ✓ {NOME_EXECUTAVEL}.exe ({tamanho:.1f} MB)")
    
    # Criar estrutura de pastas e copiar arquivos
    for pasta, arquivos in ESTRUTURA_DISTRIBUICAO.items():
        pasta_dest = os.path.join(PASTA_DIST, pasta)
        os.makedirs(pasta_dest, exist_ok=True)
        
        if arquivos:  # Se tem arquivos para copiar
            for arquivo in arquivos:
                origem = os.path.join(pasta, arquivo)
                if os.path.exists(origem):
                    destino = os.path.join(pasta_dest, arquivo)
                    shutil.copy2(origem, destino)
                    print(f"  ✓ {pasta}/{arquivo}")
                else:
                    print(f"  ⚠ {pasta}/{arquivo} não encontrado (pulando)")
        else:
            print(f"  ✓ {pasta}/ (vazia)")
    
    return True

def criar_documentacao_instalacao():
    """Cria documentação de instalação simplificada"""
    print("\n📄 Criando documentação...")
    
    # README.txt principal
    readme_content = f"""
╔══════════════════════════════════════════════════════════════════════╗
║                                                                      ║
║          ANALISADOR INSS - Sistema de Análise Previdenciária        ║
║                          Versão {VERSAO}                                 ║
║                                                                      ║
╚══════════════════════════════════════════════════════════════════════╝

📋 SOBRE
────────
Sistema para análise de elegibilidade e cálculo de aposentadoria
conforme regras pré e pós-reforma da previdência (EC 103/2019).


🚀 INSTALAÇÃO
─────────────
1. Extrair todo o conteúdo desta pasta em um local permanente
   Exemplo: C:\\Programas\\AnalisadorINSS\\

2. NÃO mover ou renomear pastas internas (fat_inpc_selic, Planilhas, etc)


▶️ USO
──────
1. Duplo-clique em: AnalisadorINSS.exe

2. Selecionar PDF do extrato CNIS

3. Aguardar processamento

4. Planilha Excel será gerada na pasta: saida\\


📊 RESULTADOS
─────────────
O sistema gera 9 abas na planilha Excel:

1. Config_Regras       - Parâmetros do sistema
2. Dados_Cliente       - Informações do segurado
3. Remuneracoes        - Histórico de contribuições corrigidas (INPC+SELIC)
4. Calculo_Media       - Cálculo da média salarial
5. Calculo_Tempo       - Tempo de contribuição
6. Calculo_Final       - Resultado consolidado
7. Analise_Pre_Reforma - Regras antigas (35H/30M + fator)
8. Analise_Pos_Reforma - Regras atuais (65H/62M + coeficiente)
9. Comparacao_Geral    - Comparação e ESTRATÉGIA ÓTIMA ⭐


⚠️ IMPORTANTE
─────────────
• As análises são baseadas em EC 103/2019 (Reforma da Previdência)
• Carência mínima: 15 anos (180 meses)
• Coeficiente 60%: 20 anos homem / 15 anos mulher
• SEMPRE validar cálculos com advogado especializado
• Este sistema é ferramenta de apoio, não substitui análise jurídica


🔧 SUPORTE TÉCNICO
──────────────────
Em caso de dúvidas ou ajustes necessários:

1. Verificar pasta info/ para documentação técnica:
   • GUIA_RAPIDO.md - Tutorial passo-a-passo
   • CALCULO_FATOR_DETALHADO.md - Fórmulas detalhadas
   • INSTALACAO_CLIENTE.md - Instruções completas

2. Contatar desenvolvedor para ajustes nas regras


📅 VERSÃO
─────────
Build: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}
Python: {sys.version.split()[0]}
Sistema: Windows 10/11 (64-bit)


═══════════════════════════════════════════════════════════════════════
"""
    
    readme_path = os.path.join(PASTA_DIST, 'LEIA-ME.txt')
    with open(readme_path, 'w', encoding='utf-8') as f:
        f.write(readme_content)
    print(f"  ✓ LEIA-ME.txt")
    
    # VALIDACAO_NECESSARIA.txt (para o advogado)
    validacao_content = """
╔══════════════════════════════════════════════════════════════════════╗
║                    ⚖️  VALIDAÇÃO JURÍDICA NECESSÁRIA                 ║
╚══════════════════════════════════════════════════════════════════════╝

Prezado Advogado,

Este sistema implementa as seguintes regras previdenciárias:

📌 PÓS-REFORMA (EC 103/2019)
────────────────────────────
✓ Carência mínima: 15 anos (180 meses) - Art. 19, EC 103/2019
✓ Idade mínima: 65 anos (homem) / 62 anos (mulher)
✓ Coeficiente base 60%: aos 20 anos (H) / 15 anos (M)
✓ Acréscimo: +2% por ano acima do tempo base
✓ Coeficiente máximo: 100% aos 35 anos (H) / 30 anos (M)
✓ Benefício mínimo: Piso INSS (R$ 1.412,00 - 2024)

📌 PRÉ-REFORMA (direito adquirido)
──────────────────────────────────
✓ Tempo contribuição: 35 anos (H) / 30 anos (M)
✓ Sem idade mínima
✓ Média: 80% maiores salários desde jul/1994
✓ Fator previdenciário: Lei 9.876/99
  • Fórmula: f = (Tc × 0,31 / Es) × [1 + (Id + Tc × 0,31) / 100]
  • Expectativa sobrevida: Tabela IBGE
  • Limitado: 0,4 a 1,3

📌 CORREÇÃO MONETÁRIA
─────────────────────
✓ Até 06/2006: Índice próprio INSS
✓ 07/2006 - 02/2025: INPC (IBGE)
✓ 03/2025 em diante: SELIC acumulada


🔍 ITENS PARA VALIDAÇÃO
═══════════════════════

[ ] 1. Carência é 15 ou 20 anos?
    Sistema usa 15 anos conforme Art. 19, EC 103/2019
    Confirmar se há interpretação divergente

[ ] 2. Coeficiente aos 15 anos
    Sistema aplica 50% (H) / 60% (M) aos 15 anos
    60% inicia aos 20H/15M
    Confirmar progressão

[ ] 3. Fator previdenciário
    Fórmula Lei 9.876/99 está correta?
    Tabela IBGE utilizada está atualizada?

[ ] 4. Piso INSS
    Valor R$ 1.412,00 (2024) - atualizar para 2026?

[ ] 5. Regras de transição
    Sistema não aplica regras de transição
    Incluir pontos (86/96, idade progressiva, pedágio)?

[ ] 6. Tempo especial
    Sistema não converte tempo especial
    Incluir conversão 1.4 (25 anos) / 1.2 (20 anos)?

[ ] 7. Contribuições facultativas
    Sistema calcula custo por salário mínimo
    Permitir simular valores diferentes?

[ ] 8. RPPS / outros vínculos
    Sistema considera apenas RGPS
    Incluir análise de tempo RPPS?

[ ] 9. Dados utilizados
    Verificar se há campos no CNIS não utilizados
    Sugerir inclusão/remoção de dados

[ ] 10. Layout das abas
    Estrutura das 9 abas está adequada?
    Sugerir reorganização/simplificação


📧 RETORNO
══════════
Após análise, favor retornar indicando:

1. ✅ Itens corretos (manter)
2. ⚠️ Itens para ajuste (especificar mudança)
3. ➕ Itens para incluir (especificar novo cálculo/dado)
4. ➖ Itens para remover (dados desnecessários)


═══════════════════════════════════════════════════════════════════════
Sistema desenvolvido como ferramenta de apoio à análise jurídica
Não substitui avaliação técnica e interpretação legal especializada
═══════════════════════════════════════════════════════════════════════
"""
    
    validacao_path = os.path.join(PASTA_DIST, 'VALIDACAO_NECESSARIA.txt')
    with open(validacao_path, 'w', encoding='utf-8') as f:
        f.write(validacao_content)
    print(f"  ✓ VALIDACAO_NECESSARIA.txt")
    
    return True

def criar_arquivo_versao():
    """Cria arquivo de versão e build info"""
    versao_path = os.path.join(PASTA_DIST, 'VERSAO.txt')
    
    with open(versao_path, 'w', encoding='utf-8') as f:
        f.write(f"AnalisadorINSS - Sistema de Análise Previdenciária\n")
        f.write(f"Versão: {VERSAO}\n")
        f.write(f"Build: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
        f.write(f"Python: {sys.version.split()[0]}\n")
        f.write(f"\nMódulos incluídos:\n")
        f.write(f"  - interface_cnis.py (GUI principal)\n")
        f.write(f"  - preparar_dados_planilha.py (geração Excel)\n")
        f.write(f"  - analise_elegibilidade.py (cálculos)\n")
        f.write(f"  - converter_extrato_inss.py (extração PDF)\n")
        f.write(f"\nDados incluídos:\n")
        f.write(f"  - INPC: 364 competências\n")
        f.write(f"  - SELIC: 474 competências\n")
        f.write(f"  - Config_Regras.xlsx: Parâmetros\n")
    
    print(f"  ✓ VERSAO.txt")

def exibir_resumo():
    """Exibe resumo do build"""
    print_header("✅ BUILD CONCLUÍDO COM SUCESSO!")
    
    print(f"\n📦 PACOTE CRIADO: {PASTA_DIST}/")
    print("\n📁 Estrutura:")
    
    # Listar conteúdo
    for root, dirs, files in os.walk(PASTA_DIST):
        level = root.replace(PASTA_DIST, '').count(os.sep)
        indent = '  ' * level
        pasta_nome = os.path.basename(root) if level > 0 else PASTA_DIST
        print(f"{indent}📁 {pasta_nome}/")
        
        sub_indent = '  ' * (level + 1)
        for arquivo in sorted(files):
            caminho = os.path.join(root, arquivo)
            tamanho = os.path.getsize(caminho)
            
            if arquivo.endswith('.exe'):
                print(f"{sub_indent}⚙️  {arquivo} ({tamanho/(1024*1024):.1f} MB)")
            else:
                if tamanho > 1024:
                    print(f"{sub_indent}📄 {arquivo} ({tamanho/1024:.0f} KB)")
                else:
                    print(f"{sub_indent}📄 {arquivo}")
    
    # Tamanho total
    tamanho_total = sum(
        os.path.getsize(os.path.join(root, file))
        for root, dirs, files in os.walk(PASTA_DIST)
        for file in files
    ) / (1024 * 1024)
    
    print(f"\n📊 Tamanho total: {tamanho_total:.1f} MB")
    
    print("\n📤 PRÓXIMOS PASSOS:")
    print("  1. Testar o executável:")
    print(f"     cd {PASTA_DIST}")
    print(f"     .\\{NOME_EXECUTAVEL}.exe")
    print("\n  2. Comprimir pasta para envio:")
    print(f"     Compress-Archive -Path {PASTA_DIST} -DestinationPath AnalisadorINSS_v{VERSAO}.zip")
    print("\n  3. Enviar ao advogado com VALIDACAO_NECESSARIA.txt")
    print("\n  4. Aguardar retorno sobre ajustes necessários")

def main():
    """Função principal do build"""
    print_header(f"🔨 BUILD PARA ADVOGADO - v{VERSAO}")
    print(f"📅 {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    
    # 1. Verificar dependências
    if not verificar_dependencias():
        return 1
    
    # 2. Limpar build anterior
    limpar_build_anterior()
    
    # 3. Criar .spec file
    spec_file = criar_spec_file()
    
    # 4. Compilar executável
    exe_path = compilar_executavel(spec_file)
    if not exe_path:
        print("\n❌ Build falhou!")
        return 1
    
    # 5. Criar estrutura de distribuição
    if not criar_estrutura_distribuicao(exe_path):
        print("\n❌ Falha ao criar estrutura!")
        return 1
    
    # 6. Criar documentação
    criar_documentacao_instalacao()
    criar_arquivo_versao()
    
    # 7. Limpar arquivos temporários
    print("\n🧹 Limpando arquivos temporários...")
    for pasta in [PASTA_BUILD, 'dist']:
        if os.path.exists(pasta):
            shutil.rmtree(pasta)
    
    for arquivo in Path('.').glob('*.spec'):
        arquivo.unlink()
    
    print("  ✓ Limpeza concluída")
    
    # 8. Exibir resumo
    exibir_resumo()
    
    return 0

if __name__ == "__main__":
    sys.exit(main())
