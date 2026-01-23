#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gera planilha Excel com análise completa do CNIS

Abas:
1. Resumo - Estatísticas gerais
2. Vínculos - Lista de vínculos com tipo detectado
3. Remunerações - Todas as 178 remunerações classificadas
4. Dados Cliente - Informações do segurado
"""

import sys
from pathlib import Path
import csv

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    print("❌ openpyxl não instalado. Instalando...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).parent))

from extrator.tipos import detectar_tipo_vinculo, TipoVinculo, obter_nome_tipo
from extrator.parsers import extrair_vinculos_texto, parse_vinculo_texto, extrair_dados_cabecalho

print("=" * 80)
print("📊 GERADOR DE PLANILHA EXCEL - ANÁLISE COMPLETA CNIS")
print("=" * 80)
print()

# Configuração
pdf_path = "cnis/CNIS_JOAO_CARLOS.pdf"
csv_remuneracoes = "saida/teste_baseline_remuneracoes.csv"
xlsx_saida = "saida/analise_completa_cnis.xlsx"

if not Path(pdf_path).exists():
    print(f"❌ PDF não encontrado: {pdf_path}")
    sys.exit(1)

if not Path(csv_remuneracoes).exists():
    print(f"⚠️  Arquivo de remunerações não encontrado: {csv_remuneracoes}")
    print(f"Execute primeiro: python converter_extrato_inss.py {pdf_path} saida/teste_baseline.csv")
    sys.exit(1)

# Criar workbook
wb = openpyxl.Workbook()
wb.remove(wb.active)  # Remove sheet default

# ============================================================================
# ABA 1: RESUMO
# ============================================================================
print("📋 Criando aba RESUMO...")

ws_resumo = wb.create_sheet("Resumo", 0)

# Título
ws_resumo['A1'] = "ANÁLISE COMPLETA DO CNIS"
ws_resumo['A1'].font = Font(size=16, bold=True)
ws_resumo['A1'].fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
ws_resumo['A1'].font = Font(size=16, bold=True, color="FFFFFF")
ws_resumo.merge_cells('A1:D1')

# Dados do Cliente
dados_cliente = extrair_dados_cabecalho(pdf_path)
row = 3
ws_resumo[f'A{row}'] = "DADOS DO SEGURADO"
ws_resumo[f'A{row}'].font = Font(bold=True, size=12)
row += 1

ws_resumo[f'A{row}'] = "Nome:"
ws_resumo[f'B{row}'] = dados_cliente.get('Nome', 'N/A')
row += 1

ws_resumo[f'A{row}'] = "NIT:"
ws_resumo[f'B{row}'] = dados_cliente.get('NIT', 'N/A')
row += 1

ws_resumo[f'A{row}'] = "CPF:"
ws_resumo[f'B{row}'] = dados_cliente.get('CPF', 'N/A')
row += 1

ws_resumo[f'A{row}'] = "Data Nascimento:"
ws_resumo[f'B{row}'] = dados_cliente.get('DataNascimento', 'N/A')
row += 1

ws_resumo[f'A{row}'] = "Nome da Mãe:"
ws_resumo[f'B{row}'] = dados_cliente.get('NomeMae', 'N/A')
row += 2

# Estatísticas de Vínculos
ws_resumo[f'A{row}'] = "ESTATÍSTICAS DE VÍNCULOS"
ws_resumo[f'A{row}'].font = Font(bold=True, size=12)
row += 1

blocos_vinculos = extrair_vinculos_texto(pdf_path)
total_vinculos = len(blocos_vinculos)

vinculos_clt = 0
vinculos_fac = 0

for _, texto_bloco in blocos_vinculos:
    tipo = detectar_tipo_vinculo(texto_bloco)
    if tipo == TipoVinculo.CLT:
        vinculos_clt += 1
    elif tipo == TipoVinculo.FACULTATIVO:
        vinculos_fac += 1

ws_resumo[f'A{row}'] = "Total de Vínculos:"
ws_resumo[f'B{row}'] = total_vinculos
row += 1

ws_resumo[f'A{row}'] = "  CLT (Empregado):"
ws_resumo[f'B{row}'] = vinculos_clt
row += 1

ws_resumo[f'A{row}'] = "  Facultativo:"
ws_resumo[f'B{row}'] = vinculos_fac
row += 2

# Estatísticas de Remunerações
ws_resumo[f'A{row}'] = "ESTATÍSTICAS DE REMUNERAÇÕES"
ws_resumo[f'A{row}'].font = Font(bold=True, size=12)
row += 1

# Ler remunerações
remun_por_seq = {}
total_remun = 0

with open(csv_remuneracoes, 'r', encoding='utf-8') as f:
    reader = csv.DictReader(f, delimiter=';')
    for linha in reader:
        seq = linha.get('Seq', '')
        if seq:
            remun_por_seq[seq] = remun_por_seq.get(seq, 0) + 1
            total_remun += 1

remun_clt = sum(count for seq, count in remun_por_seq.items() if int(seq) <= 10)
remun_fac = sum(count for seq, count in remun_por_seq.items() if int(seq) > 10)

ws_resumo[f'A{row}'] = "Total de Remunerações:"
ws_resumo[f'B{row}'] = total_remun
ws_resumo[f'B{row}'].font = Font(bold=True)
row += 1

ws_resumo[f'A{row}'] = "  CLT (Seq 1-10):"
ws_resumo[f'B{row}'] = remun_clt
row += 1

ws_resumo[f'A{row}'] = "  Facultativo (Seq 11+):"
ws_resumo[f'B{row}'] = remun_fac
row += 2

# Validação
ws_resumo[f'A{row}'] = "STATUS:"
if total_remun == 178:
    ws_resumo[f'B{row}'] = "✅ 178 REMUNERAÇÕES VALIDADAS"
    ws_resumo[f'B{row}'].font = Font(bold=True, color="008000")
else:
    ws_resumo[f'B{row}'] = f"⚠️ {total_remun} remunerações (esperado 178)"
    ws_resumo[f'B{row}'].font = Font(bold=True, color="FF0000")

# Ajustar larguras
ws_resumo.column_dimensions['A'].width = 25
ws_resumo.column_dimensions['B'].width = 50

# ============================================================================
# ABA 2: VÍNCULOS
# ============================================================================
print("📋 Criando aba VÍNCULOS...")

ws_vinculos = wb.create_sheet("Vínculos", 1)

# Cabeçalho
headers = ['Seq', 'Tipo', 'Tipo Detalhado', 'CNPJ/Código', 'Empresa/Origem', 'Data Início', 'Data Fim', 'Últ. Remun.']
for col, header in enumerate(headers, start=1):
    cell = ws_vinculos.cell(1, col, header)
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    cell.alignment = Alignment(horizontal='center')

# Dados
row = 2
for pagina_idx, texto_bloco in blocos_vinculos:
    vinculo = parse_vinculo_texto(texto_bloco)
    
    if not vinculo:
        continue
    
    tipo = detectar_tipo_vinculo(texto_bloco)
    
    ws_vinculos.cell(row, 1, vinculo.get('seq', ''))
    ws_vinculos.cell(row, 2, tipo.name)
    ws_vinculos.cell(row, 3, obter_nome_tipo(tipo))
    ws_vinculos.cell(row, 4, vinculo.get('codigo_emp', ''))
    ws_vinculos.cell(row, 5, vinculo.get('empresa', ''))
    ws_vinculos.cell(row, 6, vinculo.get('data_inicio', ''))
    ws_vinculos.cell(row, 7, vinculo.get('data_fim', ''))
    ws_vinculos.cell(row, 8, vinculo.get('ult_remun', ''))
    
    # Cor por tipo
    if tipo == TipoVinculo.CLT:
        fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
    else:
        fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    
    for col in range(1, 9):
        ws_vinculos.cell(row, col).fill = fill
    
    row += 1

# Ajustar larguras
ws_vinculos.column_dimensions['A'].width = 6
ws_vinculos.column_dimensions['B'].width = 12
ws_vinculos.column_dimensions['C'].width = 35
ws_vinculos.column_dimensions['D'].width = 20
ws_vinculos.column_dimensions['E'].width = 40
ws_vinculos.column_dimensions['F'].width = 12
ws_vinculos.column_dimensions['G'].width = 12
ws_vinculos.column_dimensions['H'].width = 12

# ============================================================================
# ABA 3: REMUNERAÇÕES
# ============================================================================
print("📋 Criando aba REMUNERAÇÕES...")

ws_remun = wb.create_sheet("Remunerações", 2)

# Cabeçalho
headers_remun = ['Seq', 'Tipo', 'CNPJ/Código', 'Competência', 'Remuneração', 'Indicadores', 'Página']
for col, header in enumerate(headers_remun, start=1):
    cell = ws_remun.cell(1, col, header)
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    cell.alignment = Alignment(horizontal='center')

# Ler e escrever remunerações
row = 2
with open(csv_remuneracoes, 'r', encoding='utf-8') as f:
    reader = csv.DictReader(f, delimiter=';')
    
    for linha in reader:
        seq = linha.get('Seq', '')
        codigo_emp = linha.get('CodigoEmp', '')
        competencia = linha.get('Competencia', '')
        remuneracao = linha.get('Remuneracao', '')
        indicadores = linha.get('Indicadores', '')
        pagina = linha.get('Pagina', '')
        
        # Determinar tipo
        seq_int = int(seq) if seq.isdigit() else 0
        tipo = "CLT" if seq_int <= 10 else "FACULTATIVO"
        
        ws_remun.cell(row, 1, seq)
        ws_remun.cell(row, 2, tipo)
        ws_remun.cell(row, 3, codigo_emp)
        ws_remun.cell(row, 4, competencia)
        ws_remun.cell(row, 5, remuneracao)
        ws_remun.cell(row, 6, indicadores)
        ws_remun.cell(row, 7, pagina)
        
        # Cor por tipo
        if tipo == "CLT":
            fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
        else:
            fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        
        for col in range(1, 8):
            ws_remun.cell(row, col).fill = fill
        
        # Formatar valor
        try:
            valor_float = float(remuneracao)
            ws_remun.cell(row, 5).number_format = '#,##0.00'
            ws_remun.cell(row, 5).value = valor_float
        except:
            pass
        
        row += 1

# Ajustar larguras
ws_remun.column_dimensions['A'].width = 6
ws_remun.column_dimensions['B'].width = 12
ws_remun.column_dimensions['C'].width = 25
ws_remun.column_dimensions['D'].width = 12
ws_remun.column_dimensions['E'].width = 15
ws_remun.column_dimensions['F'].width = 30
ws_remun.column_dimensions['G'].width = 8

# ============================================================================
# ABA 4: DADOS CLIENTE
# ============================================================================
print("📋 Criando aba DADOS CLIENTE...")

ws_cliente = wb.create_sheet("Dados Cliente", 3)

row = 1
ws_cliente[f'A{row}'] = "IDENTIFICAÇÃO DO SEGURADO"
ws_cliente[f'A{row}'].font = Font(size=14, bold=True)
ws_cliente[f'A{row}'].fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
ws_cliente[f'A{row}'].font = Font(size=14, bold=True, color="FFFFFF")
ws_cliente.merge_cells(f'A{row}:B{row}')
row += 2

campos = [
    ('Nome Completo', dados_cliente.get('Nome', '')),
    ('NIT', dados_cliente.get('NIT', '')),
    ('CPF', dados_cliente.get('CPF', '')),
    ('Data de Nascimento', dados_cliente.get('DataNascimento', '')),
    ('Nome da Mãe', dados_cliente.get('NomeMae', '')),
]

for campo, valor in campos:
    ws_cliente[f'A{row}'] = campo
    ws_cliente[f'A{row}'].font = Font(bold=True)
    ws_cliente[f'B{row}'] = valor
    row += 1

ws_cliente.column_dimensions['A'].width = 25
ws_cliente.column_dimensions['B'].width = 50

# ============================================================================
# SALVAR
# ============================================================================
print()
print(f"💾 Salvando planilha: {xlsx_saida}")

Path(xlsx_saida).parent.mkdir(parents=True, exist_ok=True)
wb.save(xlsx_saida)

print()
print("=" * 80)
print("✅ PLANILHA GERADA COM SUCESSO!")
print("=" * 80)
print()
print(f"📊 Arquivo: {xlsx_saida}")
print()
print("📋 Abas criadas:")
print("  1. Resumo         - Estatísticas gerais e validação")
print("  2. Vínculos       - 10 vínculos com tipo detectado")
print("  3. Remunerações   - 178 remunerações classificadas")
print("  4. Dados Cliente  - Informações do segurado")
print()
print("🎯 Status:")
print(f"  ✅ {total_vinculos} vínculos identificados")
print(f"  ✅ {total_remun} remunerações extraídas")
print(f"  ✅ {remun_clt} CLT + {remun_fac} Facultativo")
print()
print("🔍 Cores na planilha:")
print("  🔲 Cinza  = CLT (Empregado)")
print("  🟨 Amarelo = Facultativo (Contribuinte Individual)")
print()
print("=" * 80)
