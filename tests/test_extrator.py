"""
Testes para o extrator de CNIS.

Valida:
- Extração completa (178 remunerações)
- Distribuição correta por Seq (1-13)
- Tipos de vínculos (CLT vs Facultativo)
- Formato dos CSVs gerados
"""

import sys
import csv
from pathlib import Path

import pytest

# Adicionar diretório raiz ao path
sys.path.insert(0, str(Path(__file__).parent.parent))

from converter_extrato_inss import (
    extrair_tabelas,
    extrair_dados_cabecalho,
    salvar_raw_csv,
    salvar_cabecalho_csv,
    salvar_vinculos_brutos,
    salvar_vinculos_estruturados,
    salvar_remuneracoes_csv,
)
from extrator.tipos.coordenador_remuneracoes import extrair_remuneracoes_coordenado


# ============================================================================
# FIXTURES
# ============================================================================

@pytest.fixture
def pdf_joao_carlos():
    """Caminho para o PDF de teste (João Carlos)."""
    pdf_path = Path(__file__).parent.parent / "cnis" / "CNIS_JOAO_CARLOS.pdf"
    if not pdf_path.exists():
        pytest.skip(f"PDF de teste não encontrado: {pdf_path}")
    return str(pdf_path)


@pytest.fixture
def tmp_output_dir(tmp_path):
    """Diretório temporário para saída de testes."""
    output_dir = tmp_path / "saida"
    output_dir.mkdir()
    return output_dir


# ============================================================================
# TESTES DE EXTRAÇÃO DE REMUNERAÇÕES
# ============================================================================

def test_extracao_total_178_remuneracoes(pdf_joao_carlos):
    """Valida que extrai exatamente 178 remunerações."""
    remuneracoes = extrair_remuneracoes_coordenado(pdf_joao_carlos)
    assert len(remuneracoes) == 178, f"Esperado 178, obtido {len(remuneracoes)}"


def test_distribuicao_por_seq(pdf_joao_carlos):
    """Valida distribuição esperada por sequência (Seq 1-13)."""
    remuneracoes = extrair_remuneracoes_coordenado(pdf_joao_carlos)
    
    # Contar por Seq
    seq_counts = {}
    for r in remuneracoes:
        seq = r.get("seq", "")
        seq_counts[seq] = seq_counts.get(seq, 0) + 1
    
    # Distribuição esperada (baseada em validações anteriores)
    expected = {
        "1": 5, "2": 31, "3": 2, "4": 6, "5": 3, "6": 4,
        "7": 7, "8": 38, "9": 21, "10": 46, "11": 2, "12": 9, "13": 4
    }
    
    for seq, expected_count in expected.items():
        assert seq_counts.get(seq, 0) == expected_count, \
            f"Seq {seq}: esperado {expected_count}, obtido {seq_counts.get(seq, 0)}"


def test_vinculos_clt_tem_cnpj(pdf_joao_carlos):
    """Valida que vínculos CLT têm CNPJ (não "FACULTATIVO")."""
    remuneracoes = extrair_remuneracoes_coordenado(pdf_joao_carlos)
    
    clt = [r for r in remuneracoes if r.get("cnpj") != "FACULTATIVO"]
    
    # Todos devem ter formato CNPJ (XX.XXX.XXX/XXXX-XX)
    import re
    cnpj_pattern = re.compile(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}')
    
    for r in clt:
        cnpj = r.get("cnpj", "")
        assert cnpj_pattern.match(cnpj), f"CNPJ inválido: {cnpj}"


def test_vinculos_facultativo(pdf_joao_carlos):
    """Valida que vínculos Facultativos são identificados corretamente."""
    remuneracoes = extrair_remuneracoes_coordenado(pdf_joao_carlos)
    
    facultativos = [r for r in remuneracoes if r.get("cnpj") == "FACULTATIVO"]
    
    # Deve ter exatamente 15 (Seq 11: 2, Seq 12: 9, Seq 13: 4)
    assert len(facultativos) == 15, f"Esperado 15 facultativos, obtido {len(facultativos)}"
    
    # Todos devem ter seq >= 11
    for r in facultativos:
        seq = int(r.get("seq", "0"))
        assert seq >= 11, f"Facultativo com Seq {seq} < 11"


def test_formato_competencia(pdf_joao_carlos):
    """Valida formato de competência (MM/AAAA)."""
    remuneracoes = extrair_remuneracoes_coordenado(pdf_joao_carlos)
    
    import re
    competencia_pattern = re.compile(r'\d{2}/\d{4}')
    
    for r in remuneracoes:
        competencia = r.get("competencia", "")
        assert competencia_pattern.match(competencia), \
            f"Competência inválida: {competencia}"


def test_valores_remuneracao_validos(pdf_joao_carlos):
    """Valida que valores de remuneração são numéricos positivos."""
    remuneracoes = extrair_remuneracoes_coordenado(pdf_joao_carlos)
    
    for r in remuneracoes:
        remuneracao = r.get("remuneracao", "")
        
        # Deve ser conversível para float
        try:
            valor = float(remuneracao)
        except ValueError:
            pytest.fail(f"Remuneração não numérica: {remuneracao}")
        
        # Deve ser positivo
        assert valor > 0, f"Remuneração não positiva: {valor}"
        
        # Deve estar em faixa razoável (0 a 10 milhões)
        assert valor < 10_000_000, f"Remuneração fora da faixa: {valor}"


# ============================================================================
# TESTES DE GERAÇÃO DE CSV
# ============================================================================

def test_gerar_csv_remuneracoes(pdf_joao_carlos, tmp_output_dir):
    """Valida geração do CSV de remunerações."""
    csv_path = tmp_output_dir / "teste_remuneracoes.csv"
    
    # Extrair tabelas (necessário para salvar_remuneracoes_csv)
    linhas_saida, _ = extrair_tabelas(pdf_joao_carlos)
    
    # Gerar CSV
    salvar_remuneracoes_csv(pdf_joao_carlos, linhas_saida, str(csv_path))
    
    # Verificar existência
    assert csv_path.exists(), "CSV de remunerações não foi criado"
    
    # Verificar conteúdo
    with open(csv_path, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f, delimiter=';')
        rows = list(reader)
    
    # 178 linhas de dados
    assert len(rows) == 178, f"CSV deveria ter 178 linhas, tem {len(rows)}"
    
    # Verificar colunas esperadas
    expected_cols = {"Pagina", "Seq", "CodigoEmp", "Competencia", "Remuneracao", "Indicadores"}
    assert set(rows[0].keys()) == expected_cols, "Colunas do CSV incorretas"


def test_gerar_csv_dados_cliente(pdf_joao_carlos, tmp_output_dir):
    """Valida geração do CSV de dados do cliente."""
    csv_path = tmp_output_dir / "teste_dados_cliente.csv"
    
    # Extrair dados do cabeçalho
    dados = extrair_dados_cabecalho(pdf_joao_carlos)
    
    # Salvar CSV
    salvar_cabecalho_csv(dados, str(csv_path))
    
    # Verificar existência
    assert csv_path.exists(), "CSV de dados do cliente não foi criado"
    
    # Verificar conteúdo
    with open(csv_path, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f, delimiter=';')
        rows = list(reader)
    
    # Deve ter exatamente 1 linha
    assert len(rows) == 1, f"CSV deveria ter 1 linha, tem {len(rows)}"
    
    # Verificar campos essenciais
    row = rows[0]
    assert "Nome" in row, "Campo 'Nome' ausente"
    assert "NIT" in row, "Campo 'NIT' ausente"


def test_gerar_csv_vinculos_estruturado(pdf_joao_carlos, tmp_output_dir):
    """Valida geração do CSV de vínculos estruturados."""
    csv_path = tmp_output_dir / "teste_vinculos.csv"
    
    # Extrair tabelas e dados do cabeçalho
    linhas_saida, _ = extrair_tabelas(pdf_joao_carlos)
    dados_cab = extrair_dados_cabecalho(pdf_joao_carlos)
    
    # Salvar CSV
    salvar_vinculos_estruturados(linhas_saida, str(csv_path), dados_cab, pdf_joao_carlos)
    
    # Verificar existência
    assert csv_path.exists(), "CSV de vínculos não foi criado"
    
    # Verificar conteúdo
    with open(csv_path, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f, delimiter=';')
        rows = list(reader)
    
    # Deve ter 10 vínculos CLT (Seq 1-10)
    # Nota: Facultativos (Seq 11-13) não são extraídos como vínculos estruturados
    assert len(rows) == 10, f"CSV deveria ter 10 vínculos CLT, tem {len(rows)}"


# ============================================================================
# TESTES DE CASOS EXTREMOS
# ============================================================================

def test_pdf_inexistente():
    """Valida que PDF inexistente lança erro."""
    with pytest.raises(FileNotFoundError):
        extrair_remuneracoes_coordenado("arquivo_que_nao_existe.pdf")


def test_ordenacao_por_seq_competencia(pdf_joao_carlos):
    """Valida que remunerações são ordenadas por Seq."""
    remuneracoes = extrair_remuneracoes_coordenado(pdf_joao_carlos)
    
    # Verificar ordenação por Seq (competências podem estar fora de ordem dentro do mesmo Seq)
    for i in range(len(remuneracoes) - 1):
        seq_atual = int(remuneracoes[i].get("seq", "0"))
        seq_proximo = int(remuneracoes[i + 1].get("seq", "0"))
        
        # Seq deve estar crescente ou igual
        assert seq_atual <= seq_proximo, \
            f"Seq fora de ordem: {seq_atual} > {seq_proximo}"


# ============================================================================
# TESTES DE INTEGRAÇÃO
# ============================================================================

def test_pipeline_completo(pdf_joao_carlos, tmp_output_dir):
    """Testa pipeline completo de extração."""
    # Definir caminhos
    csv_raw = tmp_output_dir / "teste.csv"
    csv_dados = tmp_output_dir / "teste_dados_cliente.csv"
    csv_vinculos_brutos = tmp_output_dir / "teste_vinculos_brutos.csv"
    csv_vinculos = tmp_output_dir / "teste_vinculos_estruturado.csv"
    csv_remun = tmp_output_dir / "teste_remuneracoes.csv"
    
    # Extrair tabelas
    linhas_saida, max_cols = extrair_tabelas(pdf_joao_carlos)
    
    # Salvar raw
    salvar_raw_csv(linhas_saida, max_cols, str(csv_raw))
    
    # Extrair e salvar dados do cabeçalho
    dados_cab = extrair_dados_cabecalho(pdf_joao_carlos)
    salvar_cabecalho_csv(dados_cab, str(csv_dados))
    
    # Salvar vínculos brutos
    salvar_vinculos_brutos(linhas_saida, str(csv_vinculos_brutos))
    
    # Salvar vínculos estruturados
    salvar_vinculos_estruturados(linhas_saida, str(csv_vinculos), dados_cab, pdf_joao_carlos)
    
    # Salvar remunerações
    salvar_remuneracoes_csv(pdf_joao_carlos, linhas_saida, str(csv_remun))
    
    # Validar que todos os arquivos foram criados
    assert csv_raw.exists(), "CSV raw não criado"
    assert csv_dados.exists(), "CSV dados não criado"
    assert csv_vinculos_brutos.exists(), "CSV vínculos brutos não criado"
    assert csv_vinculos.exists(), "CSV vínculos não criado"
    assert csv_remun.exists(), "CSV remunerações não criado"
    
    # Validar conteúdos básicos
    with open(csv_remun, 'r', encoding='utf-8') as f:
        remun_count = sum(1 for _ in csv.DictReader(f, delimiter=';'))
    
    assert remun_count == 178, f"Esperado 178 remunerações, obtido {remun_count}"
