# Arquitetura Modular do Extrator CNIS v2.0

## 🎯 Filosofia

**Prioridade atual:** MODULARIZAÇÃO (separação clara por tipo)  
**Próxima fase:** Reutilização de código comum

Cada tipo de vínculo/contribuição tem seu próprio módulo especializado, facilitando:
- Manutenção independente
- Adição de novos tipos (MEI, Autônomo, etc)
- Testes isolados
- Documentação específica

## 📁 Estrutura

```
extrator/
├── __init__.py                    # Entry point público
│
├── tipos/                         # 🔴 Processadores por tipo de vínculo
│   ├── __init__.py
│   ├── clt.py                    # CLT (Empregado) - 3 campos
│   ├── facultativo.py            # Facultativo - 5 campos
│   ├── detector.py               # Detecta tipo de vínculo
│   └── [futuro] mei.py           # MEI
│   └── [futuro] autonomo.py      # Autônomo com CNPJ
│   └── [futuro] segurado_especial.py
│
├── parsers/                       # 🔵 Parsers especializados
│   ├── __init__.py
│   ├── vinculos.py               # Parse de vínculos (tabela + texto)
│   └── cabecalho.py              # Parse de cabeçalho (identificação)
│
├── orquestrador.py                # 🟢 Coordenador principal
├── utils.py                       # 🟡 Funções auxiliares
│
└── [legado - manter temporariamente]
    ├── extrator_clt.py           # → movido para tipos/clt.py
    ├── extrator_facultativo.py   # → movido para tipos/facultativo.py
    └── parser_vinculos.py        # → movido para parsers/vinculos.py
```

## 🔴 Módulo `tipos/` - Processadores por Tipo

### `tipos/detector.py`
**Responsabilidade:** Identificar tipo de vínculo analisando marcadores no texto.

```python
from extrator.tipos import detectar_tipo_vinculo, TipoVinculo

tipo = detectar_tipo_vinculo(texto_bloco)
if tipo == TipoVinculo.CLT:
    processar_remuneracoes_clt(...)
elif tipo == TipoVinculo.FACULTATIVO:
    processar_contribuicoes_facultativo(...)
```

**Tipos suportados:**
- `CLT` - Empregado/Agente Público (com CNPJ)
- `FACULTATIVO` - Contribuinte Individual (sem CNPJ)
- `MEI` - Microempreendedor Individual (futuro)
- `AUTONOMO` - Contribuinte com CNPJ próprio (futuro)
- `SEGURADO_ESPECIAL` - Regime especial (futuro)

### `tipos/clt.py`
**Responsabilidade:** Processar seção "Remunerações" (CLT).

**Formato:** 3 campos por competência
```
Competência | Remuneração | Indicadores
01/1995     | 286,25      | 13º SALÁRIO
```

**Características:**
- Até 3 competências por linha
- codigo_emp = CNPJ da empresa
- Regex: `(\d{2}/\d{4})\s+([\d.,]+)\s*([^\d/]*?)`

**Funções:**
- `processar_remuneracoes_clt()` - Bloco principal
- `processar_valores_soltos_clt()` - Continuação entre páginas

### `tipos/facultativo.py`
**Responsabilidade:** Processar seção "Contribuições" (Facultativo).

**Formato:** 5 campos por competência
```
Competência | Data Pagto | Contribuição | Salário | Indicadores
09/2019     | 15/09/2019 | 200,00       | 1045,00 | PREC-FACULTCONC
```

**Características:**
- Até 2 competências por linha
- codigo_emp = "FACULTATIVO" (sem CNPJ)
- Captura: Competência, Contribuição (→ remuneracao), Indicadores
- Ignora: Data Pagto, Salário
- Regex: `(\d{2}/\d{4})\s+\d{2}/\d{2}/\d{4}\s+([\d.,]+)\s+([\d.,]+)\s*([^\d/]*?)`

**Funções:**
- `processar_contribuicoes_facultativo()` - Processa seção completa

## 🔵 Módulo `parsers/` - Parsers Especializados

### `parsers/vinculos.py`
**Responsabilidade:** Extrair e parsear vínculos do PDF.

**Funções:**
- `extrair_vinculos_texto()` - Extrai blocos "Matrícula do Tipo Filiado"
- `parse_vinculo_texto()` - Parse estruturado de bloco
- `salvar_vinculos_estruturados()` - Gera CSV consolidado

### `parsers/cabecalho.py`
**Responsabilidade:** Extrair dados de identificação do filiado.

**Campos extraídos:**
- NIT
- CPF
- Nome
- Data Nascimento
- Nome da Mãe

**Funções:**
- `extrair_dados_cabecalho()` - Extrai dados
- `salvar_cabecalho_csv()` - Salva CSV

## 🟢 Módulo `orquestrador.py` - Coordenador

**Responsabilidade:** Coordenar extração completa do PDF.

**Fluxo:**
1. Extrai tabelas brutas → `[base].csv`
2. Extrai cabeçalho → `[base]_dados_cliente.csv`
3. Extrai vínculos brutos → `[base]_vinculos_brutos.csv`
4. Parse vínculos estruturados → `[base]_vinculos_estruturado.csv`
5. Processa remunerações por tipo → `[base]_remuneracoes.csv`

**Uso:**
```python
from extrator import processar_cnis_completo

arquivos = processar_cnis_completo(
    caminho_pdf="cnis/JOAO.pdf",
    pasta_saida="saida",
    nome_base="cnis_joao"
)
print(f"Remunerações: {arquivos['remuneracoes']}")
```

## 🟡 Módulo `utils.py` - Funções Auxiliares

**Funções compartilhadas:**
- `detectar_encoding()` - UTF-8 vs cp1252
- `limpar_remuneracao()` - "1.005,70" → "1005.70"
- `validar_valor()` - Range check
- `extrair_zona_util()` - Delimita conteúdo útil (entre header/footer)

## 📊 Comparação: Antes vs Depois

### Antes (Monolítico)
```
converter_extrato_inss.py (922 linhas)
├── extrair_tabelas()
├── extrair_vinculos_texto()
├── parse_vinculo_texto()
├── salvar_vinculos_estruturados()
├── extrair_dados_cabecalho()
├── processar_contribuicoes_facultativo()
├── extrair_remuneracoes_texto() (560 linhas!)
└── main()
```

### Depois (Modular)
```
converter_extrato_inss.py (~150 linhas)
└── main() → chama orquestrador

extrator/ (647 linhas distribuídas)
├── tipos/clt.py (140 linhas)
├── tipos/facultativo.py (95 linhas)
├── tipos/detector.py (135 linhas)
├── parsers/vinculos.py (218 linhas)
├── parsers/cabecalho.py (102 linhas)
├── orquestrador.py (175 linhas)
└── utils.py (160 linhas)
```

**Benefícios:**
- ✅ Responsabilidades claras
- ✅ Fácil adicionar novos tipos (MEI, Autônomo)
- ✅ Testável isoladamente
- ✅ Documentação específica por tipo
- ✅ Manutenção independente

## 🚀 Próximos Passos

### Fase 1: Integração (atual)
- [ ] Refatorar `converter_extrato_inss.py` para usar orquestrador
- [ ] Testar 178 remunerações preservadas
- [ ] Remover código duplicado (arquivos legado)

### Fase 2: Expansão de Tipos
- [ ] Implementar `tipos/mei.py`
- [ ] Implementar `tipos/autonomo.py`
- [ ] Implementar `tipos/segurado_especial.py`
- [ ] Atualizar detector para novos tipos

### Fase 3: Otimização
- [ ] Identificar código comum entre tipos
- [ ] Extrair para utils ou classe base
- [ ] Criar testes unitários
- [ ] Benchmarks de performance

## 📝 Convenções

### Nomes de Funções
- `processar_*` - Processa seção e adiciona em `registros` (mutável)
- `extrair_*` - Extrai e retorna dados (imutável)
- `salvar_*` - Persiste dados em arquivo
- `detectar_*` - Identifica/classifica dados

### Estrutura de Registros
```python
{
    "pagina": int,
    "seq": str,
    "codigo_emp": str,        # CNPJ ou "FACULTATIVO"
    "competencia": str,       # mm/aaaa
    "remuneracao": str,       # float como string
    "indicadores": str,       # marcadores
}
```

### Imports
- Relativos dentro do módulo: `from ..utils import`
- Públicos via `__init__.py`: `from extrator import`

## 🧪 Testando

```python
# Testar detector
from extrator.tipos import detectar_tipo_vinculo, TipoVinculo

texto_clt = "Código Emp. 12.345.678/0001-90"
assert detectar_tipo_vinculo(texto_clt) == TipoVinculo.CLT

texto_fac = "Origem do Vínculo\\n125.xxx RECOLHIMENTO"
assert detectar_tipo_vinculo(texto_fac) == TipoVinculo.FACULTATIVO

# Testar orquestrador
from extrator import processar_cnis_completo

arquivos = processar_cnis_completo(
    "cnis/teste.pdf",
    "saida",
    "teste"
)
assert Path(arquivos["cabecalho"]).exists()
assert Path(arquivos["vinculos_estruturado"]).exists()
```

## 📚 Referências

- [ARCHITECTURE.md](../ARCHITECTURE.md) - Arquitetura geral do sistema
- [PLANO_SIMPLIFICACAO.md](../PLANO_SIMPLIFICACAO.md) - Roadmap de refatoração
- [info/EXTRACAO_COMPLETA_ATUALIZADA.md](../info/EXTRACAO_COMPLETA_ATUALIZADA.md) - Detalhes de extração
