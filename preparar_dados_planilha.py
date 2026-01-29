"""
Preparador de Planilha de Simulação INSS - v4.0

Cria planilha Excel otimizada com:
1. Config_Regras (copiado da planilha existente)
2. Dados_Cliente (do CSV dados_cliente)
3. Remuneracoes (178 linhas com INPC + SELIC aplicados)
4. Calculo_Media (média salarial INPC + SELIC)
5. Calculo_Tempo (tempo contribuição)
6. Calculo_Final (resultado)
7. Analise_Pre_Reforma (elegibilidade regra antiga)
8. Analise_Pos_Reforma (elegibilidade regra nova)
9. Comparacao_Geral (comparação e recomendações)

Todas as fórmulas são NATIVAS do Excel (transparentes, auditáveis).
"""

import csv
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path
from datetime import datetime, date
from analise_elegibilidade import (
    calcular_idade,
    calcular_tempo_contribuicao,
    calcular_fator_previdenciario,
    calcular_media_80_maiores,
    verificar_elegibilidade_pre_reforma,
    calcular_coeficiente_pos_reforma,
    verificar_elegibilidade_pos_reforma,
    simular_contribuicao_facultativa,
    calcular_valor_para_manter_media
)

# ============================================================================
# CONFIGURAÇÕES
# ============================================================================

CSV_REMUNERACOES = "saida/teste_apos_remocao_remuneracoes.csv"
CSV_DADOS_CLIENTE = "saida/teste_apos_remocao_dados_cliente.csv"
CSV_INPC_FATORES = "fat_inpc_selic/inpc_fatores.csv"
CSV_SELIC_FATORES = "fat_inpc_selic/selic_fatores.csv"
# PLANILHA_ORIGINAL não é mais necessária - criamos do zero
PLANILHA_SAIDA = "saida/simulacao_joao_carlos.xlsx"

# ============================================================================
# ESTILOS
# ============================================================================

def criar_estilos():
    """Retorna dicionário com estilos para formatação."""
    return {
        'header': {
            'font': Font(bold=True, size=12, color='FFFFFF'),
            'fill': PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid'),
            'alignment': Alignment(horizontal='center', vertical='center')
        },
        'titulo': {
            'font': Font(bold=True, size=14, color='FFFFFF'),
            'fill': PatternFill(start_color='203864', end_color='203864', fill_type='solid'),
            'alignment': Alignment(horizontal='center', vertical='center')
        },
        'subtitulo': {
            'font': Font(bold=True, size=11),
            'fill': PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid'),
            'alignment': Alignment(horizontal='left', vertical='center')
        },
        'numero': {
            'alignment': Alignment(horizontal='right')
        },
        'moeda': {
            'alignment': Alignment(horizontal='right')
        },
        'data': {
            'alignment': Alignment(horizontal='center')
        }
    }

def aplicar_estilo(cell, estilo_dict):
    """Aplica estilo a uma célula."""
    if 'font' in estilo_dict:
        cell.font = estilo_dict['font']
    if 'fill' in estilo_dict:
        cell.fill = estilo_dict['fill']
    if 'alignment' in estilo_dict:
        cell.alignment = estilo_dict['alignment']

# ============================================================================
# CARREGAR ÍNDICES INPC E SELIC
# ============================================================================

def carregar_indices():
    """Carrega índices INPC e SELIC dos CSVs."""
    inpc = {}
    selic = {}
    
    # Carregar INPC
    try:
        with open(CSV_INPC_FATORES, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f, delimiter=';')
            for row in reader:
                comp = row['Competencia']  # MM/AAAA
                fator = float(row['Fator'])
                inpc[comp] = fator
        print(f"   ✅ INPC: {len(inpc)} competências carregadas")
    except FileNotFoundError:
        print("   ⚠️ INPC não encontrado, usando índice fixo 1.0")
    
    # Carregar SELIC
    try:
        with open(CSV_SELIC_FATORES, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f, delimiter=';')
            for row in reader:
                comp = row['Competencia']  # MM/AAAA
                fator = float(row['Fator'])
                selic[comp] = fator
        print(f"   ✅ SELIC: {len(selic)} competências carregadas")
    except FileNotFoundError:
        print("   ⚠️ SELIC não encontrado, usando índice fixo 1.0")
    
    return inpc, selic

# ============================================================================
# ABA 1: CONFIG_REGRAS
# ============================================================================

def criar_aba_config_regras(wb):
    """Cria aba Config_Regras com parâmetros do sistema."""
    print("📋 Criando aba Config_Regras...")
    
    # Criar nova aba
    ws = wb.create_sheet('Config_Regras', 0)
    
    # Título
    ws['A1'] = 'CONFIGURAÇÕES E REGRAS DO SISTEMA'
    ws.merge_cells('A1:D1')
    ws['A1'].font = Font(bold=True, size=14, color='FFFFFF')
    ws['A1'].fill = PatternFill(start_color='203864', end_color='203864', fill_type='solid')
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    
    # Cabeçalhos
    ws['A2'] = 'Parâmetro'
    ws['B2'] = 'Valor'
    ws['C2'] = 'Descrição'
    ws['D2'] = 'Base Legal'
    
    for col in ['A', 'B', 'C', 'D']:
        ws[f'{col}2'].font = Font(bold=True, color='FFFFFF')
        ws[f'{col}2'].fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        ws[f'{col}2'].alignment = Alignment(horizontal='center')
    
    # Dados
    configs = [
        ['Piso INSS 2024', 1412.00, 'Salário mínimo nacional', 'Decreto 11.864/2024'],
        ['Carência mínima', '15 anos (180 meses)', 'Tempo mínimo pós-reforma', 'EC 103/2019, Art. 19'],
        ['Idade mínima (H)', '65 anos', 'Homens pós-reforma', 'EC 103/2019, Art. 19'],
        ['Idade mínima (M)', '62 anos', 'Mulheres pós-reforma', 'EC 103/2019, Art. 19'],
        ['Coef. base', '60%', 'Aos 20 anos (H) / 15 anos (M)', 'EC 103/2019'],
        ['Acréscimo', '2% ao ano', 'Por ano acima do tempo base', 'EC 103/2019'],
        ['Coef. máximo', '100%', 'Aos 35 anos (H) / 30 anos (M)', 'EC 103/2019'],
        ['Tempo pré-reforma (H)', '35 anos', 'Direito adquirido homens', 'Art. 201, CF/88'],
        ['Tempo pré-reforma (M)', '30 anos', 'Direito adquirido mulheres', 'Art. 201, CF/88'],
        ['Correção até 06/2006', 'Índice INSS', 'Histórico próprio', 'INSS'],
        ['Correção 07/2006-02/2025', 'INPC', 'IBGE', 'INSS'],
        ['Correção 03/2025+', 'SELIC', 'STF', 'ARE 1.348.301'],
    ]
    
    linha = 3
    for config in configs:
        ws[f'A{linha}'] = config[0]
        ws[f'B{linha}'] = config[1]
        ws[f'C{linha}'] = config[2]
        ws[f'D{linha}'] = config[3]
        linha += 1
    
    # Ajustar larguras
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 50
    ws.column_dimensions['D'].width = 20
    
    print("   ✅ Config_Regras criado")

# ============================================================================
# ABA 2: DADOS_CLIENTE
# ============================================================================

def criar_aba_dados_cliente(wb):
    """Cria aba com dados do cliente (do CSV)."""
    print("📋 Criando aba Dados_Cliente...")
    
    ws = wb.create_sheet('Dados_Cliente')
    estilos = criar_estilos()
    
    # Título
    ws['A1'] = 'DADOS DO CONTRIBUINTE'
    ws.merge_cells('A1:C1')
    aplicar_estilo(ws['A1'], estilos['titulo'])
    
    # Cabeçalho
    ws['A2'] = 'Campo'
    ws['B2'] = 'Valor'
    ws['C2'] = 'Observação'
    for col in ['A2', 'B2', 'C2']:
        aplicar_estilo(ws[col], estilos['header'])
    
    # Ler dados do CSV
    try:
        with open(CSV_DADOS_CLIENTE, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f, delimiter=';')
            dados = next(reader)
        
        # Preencher dados
        linha = 3
        campos = [
            ('Nome', dados.get('Nome', '')),
            ('CPF', dados.get('CPF', '')),
            ('NIT', dados.get('NIT', '')),
            ('Data Nascimento', dados.get('DataNascimento', '')),
            ('Sexo', dados.get('Sexo', '').strip() or 'Masculino'),  # CORRIGIDO: usar do CSV ou padrão
        ]
        
        for campo, valor in campos:
            ws.cell(linha, 1, campo)
            ws.cell(linha, 2, valor)
            linha += 1
        
        print(f"   ✅ Dados do cliente: {dados.get('Nome', 'N/A')}")
        
    except FileNotFoundError:
        print("   ⚠️ CSV dados_cliente não encontrado, usando valores padrão")
        ws['A3'] = 'Nome'
        ws['B3'] = 'João Carlos Eduardo Figueiredo Basso'
    
    # Ajustar larguras
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 40
    ws.column_dimensions['C'].width = 30

# ============================================================================
# ABA 3: REMUNERACOES
# ============================================================================

def criar_aba_remuneracoes(wb, inpc_dict, selic_dict):
    """Cria aba com todas as remunerações (178 linhas) + INPC + SELIC."""
    print("📋 Criando aba Remuneracoes...")
    
    ws = wb.create_sheet('Remuneracoes')
    estilos = criar_estilos()
    
    # Título
    ws['A1'] = 'HISTÓRICO DE REMUNERAÇÕES - 178 COMPETÊNCIAS (INPC + SELIC)'
    ws.merge_cells('A1:J1')
    aplicar_estilo(ws['A1'], estilos['titulo'])
    
    # Cabeçalho
    headers = ['Pagina', 'Seq', 'CodigoEmp', 'Competencia', 'Remuneracao', 'Indicadores', 
               'Indice_INPC', 'Valor_INPC', 'Indice_SELIC', 'Valor_SELIC']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(2, col, header)
        aplicar_estilo(cell, estilos['header'])
    
    # Ler remunerações do CSV
    try:
        # CORRIGIDO: Ler e ordenar por Seq numérico
        remuneracoes = []
        with open(CSV_REMUNERACOES, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f, delimiter=';')
            for row in reader:
                remuneracoes.append(row)
        
        # Ordenar por Seq (numérico) e depois Competencia
        remuneracoes.sort(key=lambda x: (int(x.get('Seq', '0')), x.get('Competencia', '')))
        
        linha = 3
        for row in remuneracoes:
            ws.cell(linha, 1, row.get('Pagina', ''))
            ws.cell(linha, 2, int(row.get('Seq', '0')))  # Seq como número
            ws.cell(linha, 3, row.get('CodigoEmp', ''))
            ws.cell(linha, 4, row.get('Competencia', ''))
            
            # Remuneração como número
            remun = row.get('Remuneracao', '0')
            try:
                ws.cell(linha, 5, float(remun))
                ws.cell(linha, 5).number_format = '#,##0.00'
            except ValueError:
                ws.cell(linha, 5, remun)
            
            ws.cell(linha, 6, row.get('Indicadores', ''))
            
            # Buscar índices reais por competência
            competencia = row.get('Competencia', '')
            indice_inpc = inpc_dict.get(competencia, 1.0)
            indice_selic = selic_dict.get(competencia, 1.0)
            
            # Coluna G: Índice INPC
            ws.cell(linha, 7, indice_inpc)
            ws.cell(linha, 7).number_format = '0.000000'
            
            # Coluna H: Valor corrigido INPC (fórmula: E * G)
            ws.cell(linha, 8, f"=E{linha}*G{linha}")
            ws.cell(linha, 8).number_format = '#,##0.00'
            
            # Coluna I: Índice SELIC
            ws.cell(linha, 9, indice_selic)
            ws.cell(linha, 9).number_format = '0.000000'
            
            # Coluna J: Valor corrigido SELIC (fórmula: E * I)
            ws.cell(linha, 10, f"=E{linha}*I{linha}")
            ws.cell(linha, 10).number_format = '#,##0.00'
            
            linha += 1
        
        total_linhas = linha - 3
        print(f"   ✅ {total_linhas} remunerações importadas (ordenadas por Seq)")
        
    except FileNotFoundError:
        print("   ⚠️ CSV remunerações não encontrado")
        ws['A3'] = 'Erro: CSV não encontrado'
    
    # Ajustar larguras
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 6
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['E'].width = 15
    ws.column_dimensions['F'].width = 20
    ws.column_dimensions['G'].width = 15
    ws.column_dimensions['H'].width = 15
    ws.column_dimensions['I'].width = 15
    ws.column_dimensions['J'].width = 15

# ============================================================================
# ABA 4: CALCULO_MEDIA
# ============================================================================

def criar_aba_calculo_media(wb):
    """Cria aba com cálculo da média salarial (INPC e SELIC)."""
    print("📋 Criando aba Calculo_Media...")
    
    ws = wb.create_sheet('Calculo_Media')
    estilos = criar_estilos()
    
    # Título
    ws['A1'] = 'CÁLCULO DA MÉDIA SALARIAL CORRIGIDA'
    ws.merge_cells('A1:D1')
    aplicar_estilo(ws['A1'], estilos['titulo'])
    
    # Explicação
    ws['A2'] = 'Média = Soma de todos os salários corrigidos / Quantidade de meses'
    ws.merge_cells('A2:D2')
    aplicar_estilo(ws['A2'], estilos['subtitulo'])
    
    # Quantidade de meses (comum para ambos)
    # CORRIGIDO: Usar funções em inglês para compatibilidade universal
    ws['A4'] = 'Quantidade de meses'
    ws['B4'] = '=COUNT(Remuneracoes!H3:H200)'
    ws['B4'].number_format = '0'
    
    # ========== MÉDIA COM INPC ==========
    ws['A6'] = 'MÉDIA COM INPC (OFICIAL)'
    ws.merge_cells('A6:D6')
    aplicar_estilo(ws['A6'], estilos['subtitulo'])
    
    ws['A7'] = 'Soma dos salários (INPC)'
    ws['B7'] = '=SUM(Remuneracoes!H3:H200)'
    ws['B7'].number_format = '#,##0.00'
    
    ws['A8'] = 'Média salarial INPC'
    ws['B8'] = '=B7/B4'
    ws['B8'].number_format = '#,##0.00'
    aplicar_estilo(ws['A8'], estilos['subtitulo'])
    aplicar_estilo(ws['B8'], estilos['subtitulo'])
    
    # ========== MÉDIA COM SELIC ==========
    ws['A10'] = 'MÉDIA COM SELIC (EXPERIMENTAL - VALIDAR)'
    ws.merge_cells('A10:D10')
    aplicar_estilo(ws['A10'], estilos['subtitulo'])
    
    ws['A11'] = 'Soma dos salários (SELIC)'
    ws['B11'] = '=SUM(Remuneracoes!J3:J200)'
    ws['B11'].number_format = '#,##0.00'
    
    ws['A12'] = 'Média salarial SELIC'
    ws['B12'] = '=B11/B4'
    ws['B12'].number_format = '#,##0.00'
    aplicar_estilo(ws['A12'], estilos['subtitulo'])
    aplicar_estilo(ws['B12'], estilos['subtitulo'])
    
    # Nota de aviso
    ws['A13'] = 'ATENÇÃO: Fatores SELIC em escala diferente do INPC - validar com advogado'
    ws['A13'].font = Font(size=9, italic=True, color='FF0000')
    
    # ========== DIFERENÇA ==========
    ws['A14'] = 'Diferença (SELIC - INPC)'
    ws['B14'] = '=B12-B8'
    ws['B14'].number_format = '#,##0.00'
    
    ws['A15'] = 'Diferença percentual'
    ws['B15'] = '=IF(B8>0,(B12-B8)/B8,0)'
    ws['B15'].number_format = '0.00%'
    
    # Ajustar larguras
    ws.column_dimensions['A'].width = 35
    ws.column_dimensions['B'].width = 20
    
    print("   ✅ Fórmulas de média configuradas (INPC + SELIC)")

# ============================================================================
# ABA 5: CALCULO_TEMPO
# ============================================================================

def criar_aba_calculo_tempo(wb):
    """Cria aba com cálculo do tempo de contribuição."""
    print("📋 Criando aba Calculo_Tempo...")
    
    ws = wb.create_sheet('Calculo_Tempo')
    estilos = criar_estilos()
    
    # Título
    ws['A1'] = 'CÁLCULO DO TEMPO DE CONTRIBUIÇÃO'
    ws.merge_cells('A1:D1')
    aplicar_estilo(ws['A1'], estilos['titulo'])
    
    # Explicação
    ws['A2'] = 'Simplificado: 1 mês por competência (178 meses / 12 = anos)'
    ws.merge_cells('A2:D2')
    aplicar_estilo(ws['A2'], estilos['subtitulo'])
    
    # Campos
    ws['A4'] = 'Total de meses (competências)'
    ws['B4'] = '=COUNT(Remuneracoes!D3:D200)'
    ws['B4'].number_format = '0'
    
    ws['A5'] = 'Total de ANOS'
    ws['B5'] = '=B4/12'
    ws['B5'].number_format = '0.00'
    
    ws['A6'] = 'Anos acima do mínimo'
    ws['C6'] = 'Sexo:'
    ws['D6'] = '=Dados_Cliente!B6'  # Buscar sexo
    ws['B6'] = '=B5-IF(D6="Masculino",20,15)'
    ws['B6'].number_format = '0.00'
    
    # Ajustar larguras
    ws.column_dimensions['A'].width = 35
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 15
    
    print("   ✅ Fórmulas de tempo configuradas")

# ============================================================================
# ABA 6: CALCULO_FINAL
# ============================================================================

def criar_aba_calculo_final(wb):
    """Cria aba com cálculo final da aposentadoria."""
    print("📋 Criando aba Calculo_Final...")
    
    ws = wb.create_sheet('Calculo_Final')
    estilos = criar_estilos()
    
    # Título
    ws['A1'] = 'CÁLCULO FINAL DA APOSENTADORIA'
    ws.merge_cells('A1:D1')
    aplicar_estilo(ws['A1'], estilos['titulo'])
    
    # Seção 1: Dados base
    ws['A3'] = 'DADOS BASE'
    ws.merge_cells('A3:D3')
    aplicar_estilo(ws['A3'], estilos['subtitulo'])
    
    ws['A4'] = 'Média salarial corrigida (INPC)'
    ws['B4'] = '=Calculo_Media!B8'
    ws['B4'].number_format = '#,##0.00'
    
    ws['A5'] = 'Anos totais de contribuição'
    ws['B5'] = '=Calculo_Tempo!B5'
    ws['B5'].number_format = '0.00'
    
    ws['A6'] = 'Anos acima do mínimo'
    ws['B6'] = '=Calculo_Tempo!B6'
    ws['B6'].number_format = '0.00'
    
    # Seção 2: Coeficiente
    ws['A8'] = 'COEFICIENTE (60% + 2% por ano acima)'
    ws.merge_cells('A8:D8')
    aplicar_estilo(ws['A8'], estilos['subtitulo'])
    
    ws['A9'] = 'Coeficiente base (60%)'
    ws['B9'] = 0.60
    ws['B9'].number_format = '0%'
    
    ws['A10'] = 'Acréscimo (2% × anos acima)'
    ws['B10'] = '=0.02*B6'
    ws['B10'].number_format = '0.00%'
    
    ws['A11'] = 'Coeficiente total'
    ws['B11'] = '=B9+B10'
    ws['B11'].number_format = '0.00%'
    
    ws['A12'] = 'Coeficiente limitado (máx 100%)'
    ws['B12'] = '=MIN(B11,1)'
    ws['B12'].number_format = '0.00%'
    
    # Seção 3: Resultado final
    ws['A14'] = 'RESULTADO FINAL'
    ws.merge_cells('A14:D14')
    aplicar_estilo(ws['A14'], estilos['subtitulo'])
    
    ws['A15'] = 'Aposentadoria calculada'
    ws['B15'] = '=B4*B12'
    ws['B15'].number_format = '#,##0.00'
    
    ws['A16'] = 'Salário mínimo (piso)'
    ws['B16'] = '=Config_Regras!B18'
    ws['B16'].number_format = '#,##0.00'
    
    ws['A17'] = 'Teto INSS'
    ws['B17'] = '=Config_Regras!B19'
    ws['B17'].number_format = '#,##0.00'
    
    ws['A19'] = 'APOSENTADORIA FINAL'
    ws['B19'] = '=MAX(B16,MIN(B15,B17))'
    ws['B19'].number_format = '#,##0.00'
    aplicar_estilo(ws['A19'], estilos['titulo'])
    aplicar_estilo(ws['B19'], estilos['titulo'])
    
    # Ajustar larguras
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 20
    
    print("   ✅ Fórmulas finais configuradas")


# ============================================================================
# ANÁLISE PRÉ-REFORMA (EC 103/2019)
# ============================================================================

def criar_aba_analise_pre_reforma(wb):
    """
    Cria aba de análise pela regra PRÉ-REFORMA (até 13/11/2019).
    
    Regra antiga:
    - 35 anos homem / 30 anos mulher
    - Sem idade mínima
    - Fator previdenciário
    - Média dos 80% maiores salários desde 07/1994
    """
    print("   📝 Criando Analise_Pre_Reforma...")
    
    ws = wb.create_sheet("Analise_Pre_Reforma")
    estilos = criar_estilos()
    
    # TÍTULO
    ws['A1'] = 'ANÁLISE PRÉ-REFORMA (Regra Antiga)'
    ws.merge_cells('A1:F1')
    aplicar_estilo(ws['A1'], estilos['titulo'])
    
    ws['A2'] = 'Aposentadoria por Tempo de Contribuição (antes EC 103/2019)'
    ws.merge_cells('A2:F2')
    ws['A2'].font = Font(size=11, italic=True, color='666666')
    ws['A2'].alignment = Alignment(horizontal='center')
    
    # SEÇÃO 1: Dados Base
    row = 4
    ws[f'A{row}'] = 'DADOS DO SEGURADO'
    ws.merge_cells(f'A{row}:F{row}')
    aplicar_estilo(ws[f'A{row}'], estilos['subtitulo'])
    
    row += 1
    ws[f'A{row}'] = 'Nome:'
    ws[f'B{row}'] = '=Dados_Cliente!B3'
    ws[f'B{row}'].font = Font(bold=True)
    
    row += 1
    ws[f'A{row}'] = 'Sexo:'
    ws[f'B{row}'] = '=Dados_Cliente!B7'
    
    row += 1
    ws[f'A{row}'] = 'Data de Nascimento:'
    ws[f'B{row}'] = '=Dados_Cliente!B6'
    ws[f'B{row}'].number_format = 'DD/MM/YYYY'
    
    # Cálculo de idade
    dados_cliente = ler_dados_cliente()
    if dados_cliente:
        idade_info = calcular_idade(dados_cliente['data_nascimento'])
        
        row += 1
        ws[f'A{row}'] = 'Idade Atual:'
        ws[f'B{row}'] = f"{idade_info['anos']} anos e {idade_info['meses']} meses"
        ws[f'B{row}'].font = Font(bold=True, color='0000FF')
    
    # SEÇÃO 2: Requisitos
    row += 2
    ws[f'A{row}'] = 'REQUISITOS PRÉ-REFORMA'
    ws.merge_cells(f'A{row}:F{row}')
    aplicar_estilo(ws[f'A{row}'], estilos['subtitulo'])
    
    row += 1
    ws[f'A{row}'] = 'Tempo necessário:'
    if dados_cliente and dados_cliente['sexo'] == 'Masculino':
        ws[f'B{row}'] = '35 anos (420 meses)'
    else:
        ws[f'B{row}'] = '30 anos (360 meses)'
    
    row += 1
    ws[f'A{row}'] = 'Idade mínima:'
    ws[f'B{row}'] = 'Não exigida (regra antiga)'
    
    row += 1
    ws[f'A{row}'] = 'Tempo atual:'
    ws[f'B{row}'] = '=Calculo_Tempo!B7'
    ws[f'C{row}'] = 'meses'
    ws[f'D{row}'] = '=B{0}/12'.format(row)
    ws[f'E{row}'] = 'anos'
    ws[f'D{row}'].number_format = '0.00'
    
    # SEÇÃO 3: Elegibilidade
    row += 2
    ws[f'A{row}'] = 'PODE APOSENTAR HOJE?'
    ws.merge_cells(f'A{row}:F{row}')
    aplicar_estilo(ws[f'A{row}'], estilos['subtitulo'])
    
    # Calcular elegibilidade
    tempo_contrib_meses = contar_competencias()
    if dados_cliente:
        elegibilidade = verificar_elegibilidade_pre_reforma(
            tempo_contrib_meses,
            idade_info['anos'] + idade_info['meses']/12,
            dados_cliente['sexo']
        )
        
        row += 1
        ws[f'A{row}'] = 'Resposta:'
        if elegibilidade['elegivel']:
            ws[f'B{row}'] = 'SIM ✓'
            ws[f'B{row}'].font = Font(bold=True, size=14, color='008000')
        else:
            ws[f'B{row}'] = 'NÃO ✗'
            ws[f'B{row}'].font = Font(bold=True, size=14, color='FF0000')
        
        row += 1
        if not elegibilidade['elegivel']:
            ws[f'A{row}'] = 'Faltam:'
            ws[f'B{row}'] = f"{elegibilidade['falta_meses']} meses ({elegibilidade['falta_anos']:.2f} anos)"
            ws[f'B{row}'].font = Font(bold=True, color='FF0000')
        
        row += 1
        ws[f'A{row}'] = 'Observação:'
        ws[f'B{row}'] = elegibilidade['observacao']
        ws.merge_cells(f'B{row}:F{row}')
        ws[f'B{row}'].font = Font(italic=True, color='FF6600')
        ws[f'B{row}'].alignment = Alignment(wrap_text=True)
    
    # SEÇÃO 4: Cálculo do Benefício
    row += 2
    ws[f'A{row}'] = 'CÁLCULO DO BENEFÍCIO (PRÉ-REFORMA)'
    ws.merge_cells(f'A{row}:F{row}')
    aplicar_estilo(ws[f'A{row}'], estilos['subtitulo'])
    
    row += 1
    ws[f'A{row}'] = 'Média dos 80% maiores salários:'
    ws[f'B{row}'] = '=Calculo_Media!B8'
    ws[f'B{row}'].number_format = 'R$ #,##0.00'
    
    row += 1
    ws[f'A{row}'] = 'Fator Previdenciário:'
    if dados_cliente and elegibilidade:
        fator = elegibilidade['fator_previdenciario']
        ws[f'B{row}'] = fator
        ws[f'B{row}'].number_format = '0.0000'
        
        # Cor baseada no fator
        if fator >= 1.0:
            ws[f'B{row}'].font = Font(bold=True, color='008000')
        elif fator >= 0.7:
            ws[f'B{row}'].font = Font(bold=True, color='FF6600')
        else:
            ws[f'B{row}'].font = Font(bold=True, color='FF0000')
    
    row += 1
    ws[f'A{row}'] = 'Benefício (Média × Fator):'
    ws[f'B{row}'] = f'=B{row-2}*B{row-1}'
    ws[f'B{row}'].number_format = 'R$ #,##0.00'
    ws[f'B{row}'].font = Font(bold=True, size=12, color='0000FF')
    
    # SEÇÃO 5: Simulações
    row += 2
    ws[f'A{row}'] = 'SIMULAÇÕES (Contribuição Facultativa)'
    ws.merge_cells(f'A{row}:F{row}')
    aplicar_estilo(ws[f'A{row}'], estilos['subtitulo'])
    
    if dados_cliente and not elegibilidade['elegivel']:
        # Simulação 1: Salário mínimo
        row += 1
        ws[f'A{row}'] = 'Opção 1: Contribuir com salário mínimo (R$ 1.412,00)'
        ws.merge_cells(f'A{row}:F{row}')
        ws[f'A{row}'].font = Font(bold=True, color='0000FF')
        
        sim1 = simular_contribuicao_facultativa(
            tempo_contrib_meses,
            elegibilidade['tempo_atual_anos'] * 12,  # média simplificada
            1412.00,
            elegibilidade['falta_meses']
        )
        
        row += 1
        ws[f'A{row}'] = f"  → Contribuir por {elegibilidade['falta_meses']} meses"
        row += 1
        ws[f'A{row}'] = f"  → Custo total: R$ {sim1['custo_total']:,.2f}"
        row += 1
        ws[f'A{row}'] = f"  → Previsão: {sim1['data_prevista']}"
        
        # Simulação 2: Média atual
        row += 2
        media_atual = elegibilidade['tempo_atual_anos'] * 12  # simplificado
        ws[f'A{row}'] = f'Opção 2: Contribuir com média atual (~R$ 3.000,00)'
        ws.merge_cells(f'A{row}:F{row}')
        ws[f'A{row}'].font = Font(bold=True, color='0000FF')
        
        sim2 = simular_contribuicao_facultativa(
            tempo_contrib_meses,
            media_atual,
            3000.00,
            elegibilidade['falta_meses']
        )
        
        row += 1
        ws[f'A{row}'] = f"  → Contribuir por {elegibilidade['falta_meses']} meses"
        row += 1
        ws[f'A{row}'] = f"  → Custo total: R$ {sim2['custo_total']:,.2f}"
        row += 1
        ws[f'A{row}'] = f"  → Previsão: {sim2['data_prevista']}"
    
    # Ajustar larguras
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 10
    ws.column_dimensions['F'].width = 20
    
    print("   ✅ Análise pré-reforma concluída")


# ============================================================================
# ANÁLISE PÓS-REFORMA (EC 103/2019)
# ============================================================================

def criar_aba_analise_pos_reforma(wb):
    """
    Cria aba de análise pela regra PÓS-REFORMA (EC 103/2019).
    
    Regra nova:
    - 65 anos homem / 62 anos mulher (idade mínima)
    - 20 anos contribuição (mínimo)
    - Coeficiente: 60% + 2% por ano acima de 20H/15M
    - Média de 100% dos salários desde 07/1994
    """
    print("   📝 Criando Analise_Pos_Reforma...")
    
    ws = wb.create_sheet("Analise_Pos_Reforma")
    estilos = criar_estilos()
    
    # TÍTULO
    ws['A1'] = 'ANÁLISE PÓS-REFORMA (Regra Nova)'
    ws.merge_cells('A1:F1')
    aplicar_estilo(ws['A1'], estilos['titulo'])
    
    ws['A2'] = 'Aposentadoria por Idade (EC 103/2019 - 13/11/2019)'
    ws.merge_cells('A2:F2')
    ws['A2'].font = Font(size=11, italic=True, color='666666')
    ws['A2'].alignment = Alignment(horizontal='center')
    
    # SEÇÃO 1: Dados Base
    row = 4
    ws[f'A{row}'] = 'DADOS DO SEGURADO'
    ws.merge_cells(f'A{row}:F{row}')
    aplicar_estilo(ws[f'A{row}'], estilos['subtitulo'])
    
    row += 1
    ws[f'A{row}'] = 'Nome:'
    ws[f'B{row}'] = '=Dados_Cliente!B3'
    ws[f'B{row}'].font = Font(bold=True)
    
    row += 1
    ws[f'A{row}'] = 'Sexo:'
    ws[f'B{row}'] = '=Dados_Cliente!B7'
    
    row += 1
    ws[f'A{row}'] = 'Data de Nascimento:'
    ws[f'B{row}'] = '=Dados_Cliente!B6'
    ws[f'B{row}'].number_format = 'DD/MM/YYYY'
    
    # Cálculo de idade
    dados_cliente = ler_dados_cliente()
    if dados_cliente:
        idade_info = calcular_idade(dados_cliente['data_nascimento'])
        
        row += 1
        ws[f'A{row}'] = 'Idade Atual:'
        ws[f'B{row}'] = f"{idade_info['anos']} anos e {idade_info['meses']} meses"
        ws[f'B{row}'].font = Font(bold=True, color='0000FF')
    
    # SEÇÃO 2: Requisitos
    row += 2
    ws[f'A{row}'] = 'REQUISITOS PÓS-REFORMA'
    ws.merge_cells(f'A{row}:F{row}')
    aplicar_estilo(ws[f'A{row}'], estilos['subtitulo'])
    
    row += 1
    ws[f'A{row}'] = 'Idade mínima:'
    if dados_cliente and dados_cliente['sexo'] == 'Masculino':
        idade_min = 65
        ws[f'B{row}'] = '65 anos (homem)'
    else:
        idade_min = 62
        ws[f'B{row}'] = '62 anos (mulher)'
    
    row += 1
    ws[f'A{row}'] = 'Tempo mínimo (carência):'
    ws[f'B{row}'] = '15 anos (180 meses)'
    
    row += 1
    ws[f'A{row}'] = 'Tempo para coef. 60%:'
    ws[f'B{row}'] = '20 anos homem / 15 anos mulher'
    
    row += 1
    ws[f'A{row}'] = 'Idade atual:'
    ws[f'B{row}'] = f"{idade_info['anos']} anos"
    if idade_info['anos'] >= idade_min:
        ws[f'C{row}'] = '✓ OK'
        ws[f'C{row}'].font = Font(bold=True, color='008000')
    else:
        falta_idade = idade_min - idade_info['anos']
        ws[f'C{row}'] = f'✗ Faltam {falta_idade} anos'
        ws[f'C{row}'].font = Font(bold=True, color='FF0000')
    
    row += 1
    ws[f'A{row}'] = 'Tempo atual:'
    ws[f'B{row}'] = '=Calculo_Tempo!B7'
    ws[f'C{row}'] = 'meses'
    ws[f'D{row}'] = '=B{0}/12'.format(row)
    ws[f'E{row}'] = 'anos'
    ws[f'D{row}'].number_format = '0.00'
    
    # SEÇÃO 3: Elegibilidade
    row += 2
    ws[f'A{row}'] = 'PODE APOSENTAR HOJE?'
    ws.merge_cells(f'A{row}:F{row}')
    aplicar_estilo(ws[f'A{row}'], estilos['subtitulo'])
    
    # Calcular elegibilidade
    tempo_contrib_meses = contar_competencias()
    if dados_cliente:
        elegibilidade = verificar_elegibilidade_pos_reforma(
            tempo_contrib_meses,
            idade_info['anos'] + idade_info['meses']/12,
            dados_cliente['sexo']
        )
        
        row += 1
        ws[f'A{row}'] = 'Resposta:'
        if elegibilidade['elegivel']:
            ws[f'B{row}'] = 'SIM ✓'
            ws[f'B{row}'].font = Font(bold=True, size=14, color='008000')
        else:
            ws[f'B{row}'] = 'NÃO ✗'
            ws[f'B{row}'].font = Font(bold=True, size=14, color='FF0000')
        
        row += 1
        if not elegibilidade['atende_idade']:
            ws[f'A{row}'] = 'Idade:'
            ws[f'B{row}'] = f"Faltam {elegibilidade['falta_idade_anos']:.1f} anos"
            ws[f'B{row}'].font = Font(color='FF0000')
        else:
            ws[f'A{row}'] = 'Idade:'
            ws[f'B{row}'] = 'OK ✓'
            ws[f'B{row}'].font = Font(color='008000')
        
        row += 1
        if not elegibilidade['atende_tempo']:
            ws[f'A{row}'] = 'Tempo:'
            ws[f'B{row}'] = f"Faltam {elegibilidade['falta_tempo_meses']} meses"
            ws[f'B{row}'].font = Font(color='FF0000')
        else:
            ws[f'A{row}'] = 'Tempo:'
            ws[f'B{row}'] = 'OK ✓'
            ws[f'B{row}'].font = Font(color='008000')
    
    # SEÇÃO 4: Cálculo do Benefício
    row += 2
    ws[f'A{row}'] = 'CÁLCULO DO BENEFÍCIO (PÓS-REFORMA)'
    ws.merge_cells(f'A{row}:F{row}')
    aplicar_estilo(ws[f'A{row}'], estilos['subtitulo'])
    
    row += 1
    ws[f'A{row}'] = 'Média de 100% dos salários:'
    ws[f'B{row}'] = '=Calculo_Media!B8'
    ws[f'B{row}'].number_format = 'R$ #,##0.00'
    
    row += 1
    ws[f'A{row}'] = 'Coeficiente atual:'
    if dados_cliente and elegibilidade:
        coef = elegibilidade['coeficiente']
        ws[f'B{row}'] = coef
        ws[f'B{row}'].number_format = '0.00%'
        ws[f'C{row}'] = f"({elegibilidade['coeficiente_percentual']:.0f}%)"
        
        # Cor baseada no coeficiente
        if coef >= 1.0:
            ws[f'B{row}'].font = Font(bold=True, color='008000')
        elif coef >= 0.8:
            ws[f'B{row}'].font = Font(bold=True, color='FF6600')
        else:
            ws[f'B{row}'].font = Font(bold=True, color='FF0000')
    
    row += 1
    ws[f'A{row}'] = 'Benefício (Média × Coeficiente):'
    ws[f'B{row}'] = f'=B{row-2}*B{row-1}'
    ws[f'B{row}'].number_format = 'R$ #,##0.00'
    ws[f'B{row}'].font = Font(bold=True, size=12, color='0000FF')
    
    # Comparação com piso
    row += 1
    ws[f'A{row}'] = 'Piso INSS (salário mínimo):'
    ws[f'B{row}'] = 1412.00
    ws[f'B{row}'].number_format = 'R$ #,##0.00'
    ws[f'B{row}'].font = Font(color='666666')
    
    row += 1
    ws[f'A{row}'] = 'Benefício REAL a receber:'
    ws[f'B{row}'] = f'=MAX(B{row-2},B{row-1})'
    ws[f'B{row}'].number_format = 'R$ #,##0.00'
    ws[f'B{row}'].font = Font(bold=True, size=13, color='008000')
    ws.merge_cells(f'B{row}:C{row}')
    
    row += 1
    ws[f'A{row}'] = 'Observação:'
    if dados_cliente:
        beneficio_calc = elegibilidade['coeficiente'] * 1982.65  # média aproximada
        if beneficio_calc < 1412.00:
            ws[f'B{row}'] = 'INSS paga o piso (salário mínimo) pois benefício calculado é menor'
            ws[f'B{row}'].font = Font(italic=True, color='008000')
        else:
            ws[f'B{row}'] = 'Benefício calculado acima do piso'
            ws[f'B{row}'].font = Font(italic=True, color='0000FF')
        ws.merge_cells(f'B{row}:F{row}')
        ws[f'B{row}'].alignment = Alignment(wrap_text=True)
    
    row += 1
    ws[f'A{row}'] = 'Para coeficiente 100%:'
    if dados_cliente:
        falta_100 = elegibilidade['falta_100_meses']
        if falta_100 > 0:
            ws[f'B{row}'] = f"Faltam {falta_100} meses ({falta_100/12:.1f} anos)"
            ws[f'B{row}'].font = Font(italic=True, color='FF6600')
        else:
            ws[f'B{row}'] = 'Já atingiu 100% ✓'
            ws[f'B{row}'].font = Font(italic=True, color='008000')
    
    # SEÇÃO 5: Cenários
    row += 2
    ws[f'A{row}'] = 'CENÁRIOS POSSÍVEIS'
    ws.merge_cells(f'A{row}:F{row}')
    aplicar_estilo(ws[f'A{row}'], estilos['subtitulo'])
    
    if dados_cliente:
        # Cenário A: Aposentar hoje
        row += 1
        ws[f'A{row}'] = 'Cenário A: Aposentar hoje'
        ws.merge_cells(f'A{row}:F{row}')
        ws[f'A{row}'].font = Font(bold=True, color='0000FF')
        
        row += 1
        if elegibilidade['elegivel']:
            ws[f'A{row}'] = f"  → Possível: SIM"
            ws[f'B{row}'] = f"Benefício: R$ (ver cálculo acima)"
        else:
            ws[f'A{row}'] = f"  → Possível: NÃO"
            motivo = []
            if not elegibilidade['atende_idade']:
                motivo.append(f"falta idade ({elegibilidade['falta_idade_anos']:.1f} anos)")
            if not elegibilidade['atende_tempo']:
                motivo.append(f"falta tempo ({elegibilidade['falta_tempo_meses']} meses)")
            ws[f'B{row}'] = " e ".join(motivo)
        
        # Cenário B: Esperar idade mínima
        if not elegibilidade['atende_idade']:
            row += 2
            ws[f'A{row}'] = 'Cenário B: Aguardar idade mínima'
            ws.merge_cells(f'A{row}:F{row}')
            ws[f'A{row}'].font = Font(bold=True, color='0000FF')
            
            row += 1
            meses_espera = int(elegibilidade['falta_idade_anos'] * 12)
            ws[f'A{row}'] = f"  → Aguardar {meses_espera} meses ({elegibilidade['falta_idade_anos']:.1f} anos)"
            row += 1
            ws[f'A{row}'] = f"  → Sem custo adicional"
        
        # Cenário C: Contribuir até 100%
        if elegibilidade['falta_100_meses'] > 0:
            row += 2
            ws[f'A{row}'] = 'Cenário C: Contribuir até coeficiente 100%'
            ws.merge_cells(f'A{row}:F{row}')
            ws[f'A{row}'].font = Font(bold=True, color='0000FF')
            
            sim3 = simular_contribuicao_facultativa(
                tempo_contrib_meses,
                3000.00,  # média simplificada
                3000.00,
                elegibilidade['falta_100_meses']
            )
            
            row += 1
            ws[f'A{row}'] = f"  → Contribuir por {elegibilidade['falta_100_meses']} meses"
            row += 1
            ws[f'A{row}'] = f"  → Custo estimado: R$ {sim3['custo_total']:,.2f}"
            row += 1
            ws[f'A{row}'] = f"  → Previsão: {sim3['data_prevista']}"
    
    # SEÇÃO 6: ESTRATÉGIA ÓTIMA (15 ANOS - CARÊNCIA)
    row += 3
    ws[f'A{row}'] = '⭐ ESTRATÉGIA ÓTIMA - TEMPO MÍNIMO (15 ANOS - CARÊNCIA)'
    ws.merge_cells(f'A{row}:F{row}')
    ws[f'A{row}'].font = Font(bold=True, size=12, color='FFFFFF')
    ws[f'A{row}'].fill = PatternFill(start_color='00B050', end_color='00B050', fill_type='solid')
    ws[f'A{row}'].alignment = Alignment(horizontal='center')
    
    if dados_cliente:
        # Tempo mínimo = 15 anos = 180 meses (CARÊNCIA)
        falta_minimo = max(0, 180 - tempo_contrib_meses)
        
        row += 1
        ws[f'A{row}'] = 'Requisitos mínimos:'
        ws[f'B{row}'] = '15 anos (carência) + 65 anos idade'
        ws[f'B{row}'].font = Font(bold=True)
        
        row += 1
        ws[f'A{row}'] = 'Faltam para 15 anos:'
        if falta_minimo > 0:
            ws[f'B{row}'] = f"{falta_minimo} meses ({falta_minimo/12:.1f} anos)"
            ws[f'B{row}'].font = Font(bold=True, color='0000FF')
        else:
            ws[f'B{row}'] = 'Já atingiu ✓'
            ws[f'B{row}'].font = Font(bold=True, color='008000')
        
        row += 1
        ws[f'A{row}'] = 'Coeficiente com 15 anos:'
        # Homem com 15 anos: menos de 60% (precisa 20 para 60%)
        if dados_cliente['sexo'] == 'Masculino':
            ws[f'B{row}'] = '50% (precisa 20 anos para 60%)'
            ws[f'B{row}'].font = Font(color='FF6600')
        else:
            ws[f'B{row}'] = '60%'
        
        row += 1
        ws[f'A{row}'] = 'Benefício calculado:'
        # Usar coeficiente atual da análise
        ws[f'B{row}'] = '=Calculo_Media!B8*Analise_Pos_Reforma!B' + str(row-10)  # Referência ao coef
        ws[f'B{row}'].number_format = 'R$ #,##0.00'
        
        row += 1
        ws[f'A{row}'] = 'Piso INSS:'
        ws[f'B{row}'] = 1412.00
        ws[f'B{row}'].number_format = 'R$ #,##0.00'
        
        row += 1
        ws[f'A{row}'] = 'Você vai receber:'
        ws[f'B{row}'] = f'=MAX(B{row-2},B{row-1})'
        ws[f'B{row}'].number_format = 'R$ #,##0.00'
        ws[f'B{row}'].font = Font(bold=True, size=14, color='008000')
        ws.merge_cells(f'B{row}:C{row}')
        
        row += 1
        ws[f'A{row}'] = 'Custo total (salário mínimo):'
        if falta_minimo > 0:
            custo_minimo = falta_minimo * 1412.00
            ws[f'B{row}'] = custo_minimo
            ws[f'B{row}'].number_format = 'R$ #,##0.00'
            ws[f'B{row}'].font = Font(bold=True, color='0000FF')
        else:
            ws[f'B{row}'] = 'Sem custo adicional'
        
        row += 1
        ws[f'A{row}'] = 'Quando se aposenta:'
        ws[f'B{row}'] = f"{65 - idade_info['anos']} anos de espera (aos 65 anos)"
        ws[f'B{row}'].font = Font(italic=True)
        
        row += 1
        ws[f'A{row}'] = 'Vantagens:'
        ws.merge_cells(f'A{row}:F{row}')
        ws[f'A{row}'].font = Font(bold=True, color='008000')
        
        row += 1
        ws[f'A{row}'] = '  ✓ Menor investimento possível'
        row += 1
        ws[f'A{row}'] = '  ✓ Recebe o piso (acima do calculado)'
        row += 1
        ws[f'A{row}'] = '  ✓ Mesma idade de aposentadoria (65 anos)'
        row += 1
        ws[f'A{row}'] = '  ✓ Não precisa trabalhar até 73 anos'
    
    # Ajustar larguras
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 10
    ws.column_dimensions['F'].width = 20
    
    print("   ✅ Análise pós-reforma concluída")


# ============================================================================
# COMPARAÇÃO GERAL
# ============================================================================

def criar_aba_comparacao_geral(wb):
    """
    Cria aba de comparação entre pré e pós-reforma com recomendações.
    """
    print("   📝 Criando Comparacao_Geral...")
    
    ws = wb.create_sheet("Comparacao_Geral")
    estilos = criar_estilos()
    
    # TÍTULO
    ws['A1'] = 'COMPARAÇÃO: PRÉ-REFORMA × PÓS-REFORMA'
    ws.merge_cells('A1:F1')
    aplicar_estilo(ws['A1'], estilos['titulo'])
    
    ws['A2'] = 'Qual estratégia é mais vantajosa?'
    ws.merge_cells('A2:F2')
    ws['A2'].font = Font(size=11, italic=True, color='666666')
    ws['A2'].alignment = Alignment(horizontal='center')
    
    # TABELA COMPARATIVA
    row = 4
    ws[f'A{row}'] = 'CRITÉRIO'
    ws[f'B{row}'] = 'PRÉ-REFORMA (Antiga)'
    ws[f'C{row}'] = 'PÓS-REFORMA (Nova)'
    ws[f'D{row}'] = 'MELHOR OPÇÃO'
    
    for col in ['A', 'B', 'C', 'D']:
        ws[f'{col}{row}'].font = Font(bold=True, color='FFFFFF')
        ws[f'{col}{row}'].fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        ws[f'{col}{row}'].alignment = Alignment(horizontal='center', vertical='center')
    
    # Linhas da tabela
    dados_cliente = ler_dados_cliente()
    tempo_contrib_meses = contar_competencias()
    
    if dados_cliente:
        idade_info = calcular_idade(dados_cliente['data_nascimento'])
        elegib_pre = verificar_elegibilidade_pre_reforma(
            tempo_contrib_meses,
            idade_info['anos'] + idade_info['meses']/12,
            dados_cliente['sexo']
        )
        elegib_pos = verificar_elegibilidade_pos_reforma(
            tempo_contrib_meses,
            idade_info['anos'] + idade_info['meses']/12,
            dados_cliente['sexo']
        )
        
        # Linha 1: Pode aposentar hoje?
        row += 1
        ws[f'A{row}'] = 'Pode aposentar hoje?'
        ws[f'B{row}'] = 'SIM ✓' if elegib_pre['elegivel'] else 'NÃO ✗'
        ws[f'C{row}'] = 'SIM ✓' if elegib_pos['elegivel'] else 'NÃO ✗'
        
        if elegib_pre['elegivel'] and not elegib_pos['elegivel']:
            ws[f'D{row}'] = 'PRÉ-REFORMA'
            ws[f'D{row}'].font = Font(bold=True, color='008000')
        elif elegib_pos['elegivel'] and not elegib_pre['elegivel']:
            ws[f'D{row}'] = 'PÓS-REFORMA'
            ws[f'D{row}'].font = Font(bold=True, color='008000')
        elif elegib_pre['elegivel'] and elegib_pos['elegivel']:
            ws[f'D{row}'] = 'Ambas'
            ws[f'D{row}'].font = Font(bold=True, color='0000FF')
        else:
            ws[f'D{row}'] = 'Nenhuma'
            ws[f'D{row}'].font = Font(color='FF0000')
        
        # Linha 2: Tempo faltante
        row += 1
        ws[f'A{row}'] = 'Tempo faltante'
        ws[f'B{row}'] = f"{elegib_pre['falta_meses']} meses" if elegib_pre['falta_meses'] > 0 else 'Completo'
        ws[f'C{row}'] = f"{elegib_pos['falta_tempo_meses']} meses" if elegib_pos['falta_tempo_meses'] > 0 else 'Completo'
        
        if elegib_pre['falta_meses'] < elegib_pos['falta_tempo_meses']:
            ws[f'D{row}'] = 'PRÉ-REFORMA'
        elif elegib_pos['falta_tempo_meses'] < elegib_pre['falta_meses']:
            ws[f'D{row}'] = 'PÓS-REFORMA'
        else:
            ws[f'D{row}'] = 'Igual'
        
        # Linha 3: Requisito idade
        row += 1
        ws[f'A{row}'] = 'Requisito de idade'
        ws[f'B{row}'] = 'Não exigido'
        idade_min = 65 if dados_cliente['sexo'] == 'Masculino' else 62
        ws[f'C{row}'] = f"{idade_min} anos"
        ws[f'D{row}'] = 'PRÉ-REFORMA'
        ws[f'D{row}'].font = Font(bold=True, color='008000')
        
        # Linha 4: Fator/Coeficiente
        row += 1
        ws[f'A{row}'] = 'Fator/Coeficiente'
        ws[f'B{row}'] = f"{elegib_pre['fator_previdenciario']:.4f}"
        ws[f'C{row}'] = f"{elegib_pos['coeficiente']:.2%}"
        
        if elegib_pre['fator_previdenciario'] > elegib_pos['coeficiente']:
            ws[f'D{row}'] = 'PRÉ-REFORMA'
        else:
            ws[f'D{row}'] = 'PÓS-REFORMA'
        
        # Linha 5: Base de cálculo
        row += 1
        ws[f'A{row}'] = 'Base de cálculo'
        ws[f'B{row}'] = '80% maiores'
        ws[f'C{row}'] = '100% todos'
        ws[f'D{row}'] = 'PRÉ-REFORMA'
        ws[f'D{row}'].font = Font(bold=True, color='008000')
        
        # Linha 6: Status legal
        row += 1
        ws[f'A{row}'] = 'Status legal'
        ws[f'B{row}'] = 'Extinta (direito adquirido)'
        ws[f'C{row}'] = 'Vigente'
        ws[f'D{row}'] = 'PÓS-REFORMA'
        ws[f'D{row}'].font = Font(bold=True, color='008000')
    
    # RECOMENDAÇÃO AUTOMÁTICA
    row += 2
    ws[f'A{row}'] = '⭐ ESTRATÉGIA MAIS VANTAJOSA'
    ws.merge_cells(f'A{row}:F{row}')
    ws[f'A{row}'].font = Font(bold=True, size=13, color='FFFFFF')
    ws[f'A{row}'].fill = PatternFill(start_color='00B050', end_color='00B050', fill_type='solid')
    ws[f'A{row}'].alignment = Alignment(horizontal='center')
    
    row += 1
    if dados_cliente:
        # Calcular tempo para 15 anos (carência mínima)
        tempo_15_anos = 180  # meses
        falta_15 = max(0, tempo_15_anos - tempo_contrib_meses)
        
        ws[f'A{row}'] = '→ APOSENTAR COM TEMPO MÍNIMO (15 ANOS - CARÊNCIA)'
        ws.merge_cells(f'A{row}:F{row}')
        ws[f'A{row}'].font = Font(bold=True, size=12, color='008000')
        
        row += 1
        ws[f'A{row}'] = 'Estratégia:'
        ws[f'B{row}'] = f"Contribuir {falta_15} meses ({falta_15/12:.1f} anos) + aguardar idade 65 anos"
        ws.merge_cells(f'B{row}:F{row}')
        ws[f'B{row}'].alignment = Alignment(wrap_text=True)
        
        row += 1
        ws[f'A{row}'] = 'Tempo total contribuição:'
        ws[f'B{row}'] = '15 anos (carência mínima - EC 103/2019)'
        ws[f'B{row}'].font = Font(bold=True)
        
        row += 1
        ws[f'A{row}'] = 'Coeficiente:'
        ws[f'B{row}'] = '60%'
        
        row += 1
        ws[f'A{row}'] = 'Benefício calculado:'
        ws[f'B{row}'] = '=Calculo_Media!B8*0.6'
        ws[f'B{row}'].number_format = 'R$ #,##0.00'
        
        row += 1
        ws[f'A{row}'] = 'Piso INSS:'
        ws[f'B{row}'] = 1412.00
        ws[f'B{row}'].number_format = 'R$ #,##0.00'
        
        row += 1
        ws[f'A{row}'] = 'Você vai RECEBER:'
        ws[f'B{row}'] = '=MAX(B' + str(row-2) + ',B' + str(row-1) + ')'
        ws[f'B{row}'].number_format = 'R$ #,##0.00'
        ws[f'B{row}'].font = Font(bold=True, size=14, color='008000')
        ws.merge_cells(f'B{row}:C{row}')
        
        row += 1
        ws[f'A{row}'] = 'Comparação benefício vs piso:'
        ws.merge_cells(f'A{row}:F{row}')
        ws[f'A{row}'].font = Font(bold=True, color='0000FF')
        
        media_approx = 1982.65  # valor aproximado
        beneficio_60 = media_approx * 0.60
        
        row += 1
        ws[f'A{row}'] = f'  • Benefício calculado (60%): R$ {beneficio_60:.2f}'
        ws.merge_cells(f'A{row}:F{row}')
        
        row += 1
        ws[f'A{row}'] = f'  • Piso INSS: R$ 1.412,00'
        ws.merge_cells(f'A{row}:F{row}')
        
        row += 1
        if beneficio_60 < 1412.00:
            ws[f'A{row}'] = f'  • Como R$ {beneficio_60:.2f} < R$ 1.412,00'
            ws.merge_cells(f'A{row}:F{row}')
            ws[f'A{row}'].font = Font(color='FF6600')
            
            row += 1
            ws[f'A{row}'] = '  → INSS paga o PISO automaticamente! ✓'
            ws.merge_cells(f'A{row}:F{row}')
            ws[f'A{row}'].font = Font(bold=True, color='008000')
        else:
            ws[f'A{row}'] = f'  • Como R$ {beneficio_60:.2f} > R$ 1.412,00'
            ws.merge_cells(f'A{row}:F{row}')
            
            row += 1
            ws[f'A{row}'] = '  → Recebe o benefício calculado ✓'
            ws.merge_cells(f'A{row}:F{row}')
            ws[f'A{row}'].font = Font(bold=True, color='008000')
        
        row += 1
        ws[f'A{row}'] = 'Custo total:'
        if falta_15 > 0:
            custo = falta_15 * 1412.00
            ws[f'B{row}'] = custo
            ws[f'B{row}'].number_format = 'R$ #,##0.00'
            ws[f'B{row}'].font = Font(bold=True, color='0000FF')
            ws[f'C{row}'] = f'({falta_15} meses × R$ 1.412,00)'
            ws[f'C{row}'].font = Font(italic=True, size=9)
        else:
            ws[f'B{row}'] = 'Sem custo adicional'
            ws[f'B{row}'].font = Font(bold=True, color='008000')
        
        row += 1
        ws[f'A{row}'] = 'Quando se aposenta:'
        anos_espera = 65 - idade_info['anos']
        ws[f'B{row}'] = f"Aos 65 anos ({anos_espera} anos de espera)"
        
        row += 2
        ws[f'A{row}'] = 'Por que é a melhor opção?'
        ws.merge_cells(f'A{row}:F{row}')
        ws[f'A{row}'].font = Font(bold=True, color='0000FF')
        
        row += 1
        ws[f'A{row}'] = '  ✓ Menor investimento necessário'
        ws.merge_cells(f'A{row}:F{row}')
        
        row += 1
        ws[f'A{row}'] = '  ✓ Recebe o piso INSS garantido'
        ws.merge_cells(f'A{row}:F{row}')
        
        row += 1
        ws[f'A{row}'] = '  ✓ Mesma idade de aposentadoria (65 anos)'
        ws.merge_cells(f'A{row}:F{row}')
        
        row += 1
        ws[f'A{row}'] = '  ✓ Não precisa trabalhar até 73 anos (35 anos contribuição)'
        ws.merge_cells(f'A{row}:F{row}')
        
        row += 1
        ws[f'A{row}'] = '  ✓ Melhor retorno sobre investimento'
        ws.merge_cells(f'A{row}:F{row}')
    
    # COMPARAÇÃO COM OUTRAS OPÇÕES
    row += 2
    ws[f'A{row}'] = 'COMPARAÇÃO: TEMPO MÍNIMO vs TEMPO MÁXIMO'
    ws.merge_cells(f'A{row}:F{row}')
    aplicar_estilo(ws[f'A{row}'], estilos['subtitulo'])
    
    row += 1
    ws[f'A{row}'] = '⚠️ VALIDAR COM ADVOGADO: Carência 15 ou 20 anos? (usando 15)'
    ws.merge_cells(f'A{row}:F{row}')
    ws[f'A{row}'].font = Font(bold=True, italic=True, color='FF6600')
    ws[f'A{row}'].fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    
    row += 1
    if dados_cliente:
        # Cabeçalho tabela
        ws[f'A{row}'] = 'CRITÉRIO'
        ws[f'B{row}'] = '15 ANOS ⭐'
        ws[f'C{row}'] = '35 ANOS'
        ws[f'D{row}'] = 'DIFERENÇA'
        
        for col in ['A', 'B', 'C', 'D']:
            ws[f'{col}{row}'].font = Font(bold=True, color='FFFFFF')
            ws[f'{col}{row}'].fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
            ws[f'{col}{row}'].alignment = Alignment(horizontal='center')
        
        # Linha 1: Tempo contribuir
        row += 1
        ws[f'A{row}'] = 'Tempo contribuir'
        falta_15 = max(0, 180 - tempo_contrib_meses)
        falta_35 = max(0, 420 - tempo_contrib_meses)
        ws[f'B{row}'] = f"{falta_15/12:.1f} anos"
        ws[f'C{row}'] = f"{falta_35/12:.1f} anos"
        ws[f'D{row}'] = f"{(falta_35-falta_15)/12:.1f} anos a mais"
        ws[f'D{row}'].font = Font(color='FF0000')
        
        # Linha 2: Custo
        row += 1
        ws[f'A{row}'] = 'Custo total'
        custo_15 = falta_15 * 1412.00
        custo_35 = falta_35 * 3000.00
        ws[f'B{row}'] = custo_15
        ws[f'B{row}'].number_format = 'R$ #,##0.00'
        ws[f'C{row}'] = custo_35
        ws[f'C{row}'].number_format = 'R$ #,##0.00'
        ws[f'D{row}'] = custo_35 - custo_15
        ws[f'D{row}'].number_format = 'R$ #,##0.00'
        ws[f'D{row}'].font = Font(color='FF0000')
        
        # Linha 3: Quando aposenta
        row += 1
        ws[f'A{row}'] = 'Quando aposenta'
        ws[f'B{row}'] = '2039 (65 anos)'
        ws[f'C{row}'] = '2046 (73 anos)'
        ws[f'D{row}'] = '7 anos a mais'
        ws[f'D{row}'].font = Font(color='FF0000')
        
        # Linha 4: Benefício
        row += 1
        ws[f'A{row}'] = 'Benefício mensal'
        ws[f'B{row}'] = 1412.00
        ws[f'B{row}'].number_format = 'R$ #,##0.00'
        ws[f'C{row}'] = 1982.65
        ws[f'C{row}'].number_format = 'R$ #,##0.00'
        diferenca_mensal = 1982.65 - 1412.00
        ws[f'D{row}'] = diferenca_mensal
        ws[f'D{row}'].number_format = 'R$ #,##0.00'
        ws[f'D{row}'].font = Font(color='008000')
        
        # Linha 5: ROI
        row += 1
        ws[f'A{row}'] = 'Meses para recuperar'
        if custo_35 > custo_15 and diferenca_mensal > 0:
            meses_recuperar = int((custo_35 - custo_15) / diferenca_mensal)
            ws[f'B{row}'] = '-'
            ws[f'C{row}'] = f"{meses_recuperar} meses"
            ws[f'D{row}'] = f"{meses_recuperar/12:.1f} anos"
            ws[f'D{row}'].font = Font(bold=True, color='FF0000')
        
        # Conclusão
        row += 2
        ws[f'A{row}'] = 'CONCLUSÃO:'
        ws.merge_cells(f'A{row}:F{row}')
        ws[f'A{row}'].font = Font(bold=True, size=11, color='0000FF')
        
        row += 1
        ws[f'A{row}'] = f'Para recuperar o investimento adicional de R$ {custo_35-custo_15:,.2f},'
        ws.merge_cells(f'A{row}:F{row}')
        
        row += 1
        if meses_recuperar > 0:
            ws[f'A{row}'] = f'precisaria receber por {meses_recuperar} meses ({meses_recuperar/12:.0f} anos).'
            ws.merge_cells(f'A{row}:F{row}')
            
            row += 1
            ws[f'A{row}'] = '→ OPÇÃO 15 ANOS É MUITO MAIS VANTAJOSA! ⭐'
            ws.merge_cells(f'A{row}:F{row}')
            ws[f'A{row}'].font = Font(bold=True, size=12, color='008000')
    
    # OUTRAS RECOMENDAÇÕES
    row += 2
    ws[f'A{row}'] = 'OUTRAS OBSERVAÇÕES'
    ws.merge_cells(f'A{row}:F{row}')
    aplicar_estilo(ws[f'A{row}'], estilos['subtitulo'])
    
    row += 1
    if dados_cliente:
        if elegib_pre['direito_adquirido']:
            ws[f'A{row}'] = '✓ Cliente tem DIREITO ADQUIRIDO à regra pré-reforma!'
            ws.merge_cells(f'A{row}:F{row}')
            ws[f'A{row}'].font = Font(bold=True, size=12, color='008000')
            ws[f'A{row}'].alignment = Alignment(wrap_text=True)
            
            row += 1
            ws[f'A{row}'] = 'Recomenda-se analisar valor do benefício pela regra antiga vs regra nova.'
            ws.merge_cells(f'A{row}:F{row}')
            ws[f'A{row}'].font = Font(italic=True)
            ws[f'A{row}'].alignment = Alignment(wrap_text=True)
        
        elif elegib_pre['elegivel'] and not elegib_pos['elegivel']:
            ws[f'A{row}'] = '→ Buscar reconhecimento de direito adquirido (regra antiga)'
            ws.merge_cells(f'A{row}:F{row}')
            ws[f'A{row}'].font = Font(bold=True, color='0000FF')
            
            row += 1
            ws[f'A{row}'] = 'Cliente completou tempo antes da reforma, mas precisa comprovar.'
            ws.merge_cells(f'A{row}:F{row}')
            ws[f'A{row}'].font = Font(italic=True)
            ws[f'A{row}'].alignment = Alignment(wrap_text=True)
        
        elif elegib_pos['elegivel'] and not elegib_pre['elegivel']:
            ws[f'A{row}'] = '→ Aposentar pela regra nova (pós-reforma)'
            ws.merge_cells(f'A{row}:F{row}')
            ws[f'A{row}'].font = Font(bold=True, color='0000FF')
            
            row += 1
            ws[f'A{row}'] = 'Cliente atende requisitos da regra atual.'
            ws.merge_cells(f'A{row}:F{row}')
            ws[f'A{row}'].font = Font(italic=True)
        
        else:
            ws[f'A{row}'] = '→ Cliente ainda não atende requisitos de nenhuma regra'
            ws.merge_cells(f'A{row}:F{row}')
            ws[f'A{row}'].font = Font(bold=True, color='FF0000')
            
            row += 1
            tempo_min_falta = min(elegib_pre['falta_meses'], elegib_pos['falta_tempo_meses'])
            ws[f'A{row}'] = f'Recomenda-se contribuir por pelo menos {tempo_min_falta} meses.'
            ws.merge_cells(f'A{row}:F{row}')
            ws[f'A{row}'].font = Font(italic=True)
            ws[f'A{row}'].alignment = Alignment(wrap_text=True)
    
    # AÇÕES NECESSÁRIAS
    row += 2
    ws[f'A{row}'] = 'AÇÕES NECESSÁRIAS'
    ws.merge_cells(f'A{row}:F{row}')
    aplicar_estilo(ws[f'A{row}'], estilos['subtitulo'])
    
    row += 1
    ws[f'A{row}'] = '☐ Verificar CNIS completo (sem lacunas)'
    ws.merge_cells(f'A{row}:F{row}')
    
    row += 1
    ws[f'A{row}'] = '☐ Conferir períodos especiais (insalubridade, periculosidade)'
    ws.merge_cells(f'A{row}:F{row}')
    
    row += 1
    ws[f'A{row}'] = '☐ Analisar regras de transição (pedágio 50%, 100%, idade progressiva)'
    ws.merge_cells(f'A{row}:F{row}')
    
    row += 1
    ws[f'A{row}'] = '☐ Simular cenários de contribuição facultativa'
    ws.merge_cells(f'A{row}:F{row}')
    
    row += 1
    ws[f'A{row}'] = '☐ Validar cálculos com advogado previdenciário'
    ws.merge_cells(f'A{row}:F{row}')
    
    # Ajustar larguras
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 25
    ws.column_dimensions['D'].width = 20
    ws.column_dimensions['E'].width = 15
    ws.column_dimensions['F'].width = 15
    
    print("   ✅ Comparação geral concluída")


# ============================================================================
# FUNÇÕES AUXILIARES PARA ANÁLISES
# ============================================================================

def ler_dados_cliente():
    """Lê dados do cliente do CSV."""
    try:
        with open(CSV_DADOS_CLIENTE, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f, delimiter=';')
            for row in reader:
                # Tratar sexo corretamente
                sexo_raw = row.get('Sexo', '').strip()
                if not sexo_raw or sexo_raw.lower() in ['', 'none', 'null']:
                    sexo_raw = 'Masculino'  # Default
                
                # Tratar data de nascimento
                data_nasc = row.get('DataNascimento', '') or row.get('Data_Nascimento', '')
                
                return {
                    'nome': row.get('Nome', ''),
                    'data_nascimento': data_nasc,
                    'sexo': sexo_raw,
                    'cpf': row.get('CPF', '')
                }
    except Exception as e:
        print(f"   ⚠️  Erro ao ler dados cliente: {e}")
        return None


def contar_competencias():
    """Conta total de competências (meses) com remunerações."""
    try:
        with open(CSV_REMUNERACOES, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f, delimiter=';')
            return sum(1 for _ in reader)
    except:
        return 0


# ============================================================================
# MAIN
# ============================================================================

def main():
    print("\n" + "="*80)
    print("   📊 PREPARADOR DE PLANILHA DE SIMULAÇÃO INSS - v3.0 (INPC + SELIC)")
    print("="*80 + "\n")
    
    # Carregar índices INPC e SELIC
    print("📊 Carregando índices de correção monetária...")
    inpc_dict, selic_dict = carregar_indices()
    print()
    
    # Criar workbook
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remover Sheet padrão
    
    # Criar abas básicas
    criar_aba_config_regras(wb)
    criar_aba_dados_cliente(wb)
    criar_aba_remuneracoes(wb, inpc_dict, selic_dict)
    criar_aba_calculo_media(wb)
    criar_aba_calculo_tempo(wb)
    criar_aba_calculo_final(wb)
    
    # Criar abas de análise (NOVAS - v4.0)
    print("🔍 Gerando análises de elegibilidade...")
    criar_aba_analise_pre_reforma(wb)
    criar_aba_analise_pos_reforma(wb)
    criar_aba_comparacao_geral(wb)
    
    # Salvar
    output_path = Path(PLANILHA_SAIDA)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(PLANILHA_SAIDA)
    
    print("\n" + "="*80)
    print(f"   ✅ PLANILHA CRIADA: {PLANILHA_SAIDA}")
    print(f"   📋 9 abas: Config, Cliente, Remuneracoes, Media, Tempo, Final")
    print(f"             + Pre_Reforma, Pos_Reforma, Comparacao")
    print("="*80)
    print("\n📋 ABAS CRIADAS:")
    print("   1. Config_Regras      → Parâmetros do sistema")
    print("   2. Dados_Cliente      → Dados do João Carlos")
    print("   3. Remuneracoes       → 178 competências com INPC")
    print("   4. Calculo_Media      → Média salarial (fórmulas)")
    print("   5. Calculo_Tempo      → Tempo contribuição (fórmulas)")
    print("   6. Calculo_Final      → Resultado final (fórmulas)")
    print("\n🎯 PRÓXIMO PASSO:")
    print("   Abrir planilha no Excel e verificar resultado em Calculo_Final!B19")
    print()

if __name__ == "__main__":
    main()
