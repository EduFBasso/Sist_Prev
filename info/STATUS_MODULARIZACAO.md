# Status da Modularização - Extrator CNIS

**Data:** 23/01/2026  
**Commit:** 43c2739  
**Fase:** Modularização por Tipos (Sessão 3 - em andamento)

## ✅ Completado

### Estrutura Modular Criada
```
extrator/
├── tipos/                    ✅ Criado
│   ├── __init__.py          ✅ 
│   ├── clt.py               ✅ 140 linhas (copiado de extrator_clt.py)
│   ├── facultativo.py       ✅ 95 linhas (copiado de extrator_facultativo.py)
│   └── detector.py          ✅ 135 linhas (NOVO - detecta tipo de vínculo)
│
├── parsers/                  ✅ Criado
│   ├── __init__.py          ✅
│   ├── vinculos.py          ✅ 218 linhas (copiado de parser_vinculos.py)
│   └── cabecalho.py         ✅ 102 linhas (NOVO - extraído de utils.py)
│
├── orquestrador.py           ✅ 175 linhas (NOVO - coordena extração)
├── utils.py                  ✅ 160 linhas (mantido)
├── __init__.py               ✅ Atualizado com nova estrutura
└── README_MODULAR.md         ✅ Documentação completa
```

### Novos Módulos

#### 1. `tipos/detector.py` (135 linhas)
**Responsabilidade:** Detectar tipo de vínculo (CLT, Facultativo, etc)

**Enum TipoVinculo:**
- CLT
- FACULTATIVO
- MEI (futuro)
- AUTONOMO (futuro)
- SEGURADO_ESPECIAL (futuro)
- DESCONHECIDO

**Funções:**
- `detectar_tipo_vinculo(texto_bloco)` - Analisa texto e retorna tipo
- `detectar_tipo_vinculo_por_seq(seq, nit, cnpj)` - Detecta por dados estruturados
- `obter_nome_tipo(tipo)` - Nome legível

**Lógica de Detecção:**
```python
# CLT: Possui "Código Emp." (CNPJ)
if "Código Emp." in texto:
    return TipoVinculo.CLT

# FACULTATIVO: "Origem do Vínculo" + NIT + "RECOLHIMENTO" sem CNPJ
if "Origem do Vínculo" in texto and NIT_PATTERN and "RECOLHIMENTO":
    return TipoVinculo.FACULTATIVO
```

#### 2. `parsers/cabecalho.py` (102 linhas)
**Responsabilidade:** Extrair dados de identificação do filiado

**Funções:**
- `extrair_dados_cabecalho(caminho_pdf)` → dict
- `salvar_cabecalho_csv(dados, caminho_csv)`

**Campos extraídos:**
- NIT, CPF, Nome, DataNascimento, NomeMae

#### 3. `orquestrador.py` (175 linhas)
**Responsabilidade:** Coordenar extração completa do PDF

**Funções:**
- `extrair_tabelas_brutas(pdf)` → (linhas, max_cols)
- `salvar_tabelas_raw(linhas, max_cols, csv)`
- `salvar_vinculos_brutos(linhas, csv)`
- `processar_cnis_completo(pdf, pasta, nome)` → dict de arquivos

**Fluxo de Processamento:**
```python
arquivos = processar_cnis_completo(
    caminho_pdf="cnis/JOAO.pdf",
    pasta_saida="saida",
    nome_base="cnis_joao"
)
# Retorna:
# {
#     "raw": "saida/cnis_joao.csv",
#     "cabecalho": "saida/cnis_joao_dados_cliente.csv",
#     "vinculos_brutos": "saida/cnis_joao_vinculos_brutos.csv",
#     "vinculos_estruturado": "saida/cnis_joao_vinculos_estruturado.csv",
#     "remuneracoes": "saida/cnis_joao_remuneracoes.csv"
# }
```

### Commits Realizados

**Commit 5c2f10d:** (anterior)
- Criou módulos base (utils, extrator_clt, extrator_facultativo, parser_vinculos)
- 647 linhas adicionadas

**Commit 43c2739:** (atual)
- Reorganizou em estrutura tipos/ e parsers/
- Criou detector de tipos (TipoVinculo enum)
- Criou orquestrador de extração
- Criou parser de cabeçalho separado
- Documentação completa (README_MODULAR.md)
- 1203 linhas adicionadas, 18 removidas

## ⏳ Pendente

### 1. Integrar no converter_extrato_inss.py
**Status:** NÃO INICIADO  
**Arquivo:** `converter_extrato_inss.py` (922 linhas - monolítico)

**Ações necessárias:**
```python
# Antes (922 linhas):
def main():
    linhas, max_cols = extrair_tabelas(pdf_in)
    salvar_raw_csv(linhas, max_cols, csv_out)
    dados_cab = extrair_dados_cabecalho(pdf_in)
    salvar_cabecalho_csv(dados_cab, ...)
    salvar_vinculos_brutos(linhas, ...)
    salvar_vinculos_estruturados(linhas, ...)
    salvar_remuneracoes_csv(pdf_in, linhas, ...)

# Depois (~150 linhas):
from extrator import processar_cnis_completo

def main():
    arquivos = processar_cnis_completo(
        caminho_pdf=pdf_in,
        pasta_saida=str(Path(csv_out).parent),
        nome_base=Path(csv_out).stem
    )
    for tipo, caminho in arquivos.items():
        print(f"{tipo}: {caminho}")
```

**Desafio:** Função `extrair_remuneracoes_texto()` (560 linhas) ainda não modularizada.

### 2. Refatorar extrair_remuneracoes_texto()
**Status:** BLOQUEIO  
**Arquivo:** `converter_extrato_inss.py` linhas 510-770

**Problema:** 
- Função monolítica com lógica CLT + Facultativo misturada
- 560 linhas de código procedural
- Dificulta uso do detector e processadores por tipo

**Solução proposta:**
```python
# Criar extrator/tipos/coordenador_remuneracoes.py
def extrair_remuneracoes_por_tipo(caminho_pdf):
    registros = []
    
    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            
            # Detectar vínculos na página
            for bloco in encontrar_blocos_vinculos(texto):
                tipo = detectar_tipo_vinculo(bloco)
                
                if tipo == TipoVinculo.CLT:
                    processar_remuneracoes_clt(bloco, ...)
                elif tipo == TipoVinculo.FACULTATIVO:
                    processar_contribuicoes_facultativo(bloco, ...)
                # ...
    
    return registros
```

### 3. Remover Arquivos Legado
**Status:** PENDENTE (após integração)

Arquivos a remover (mantidos temporariamente para fallback):
- [ ] `extrator/extrator_clt.py` → movido para `tipos/clt.py`
- [ ] `extrator/extrator_facultativo.py` → movido para `tipos/facultativo.py`
- [ ] `extrator/parser_vinculos.py` → movido para `parsers/vinculos.py`

### 4. Testar 178 Remunerações
**Status:** PENDENTE  
**Teste:** `cnis/CNIS_JOAO_CARLOS.pdf` → 178 linhas (163 CLT + 15 Facultativo)

**Comando:**
```bash
python converter_extrato_inss.py cnis/CNIS_JOAO_CARLOS.pdf saida/teste.csv
wc -l saida/teste_remuneracoes.csv  # Deve ser 179 (178 + header)
```

## 📊 Métricas

### Linhas de Código

| Módulo | Linhas | Descrição |
|--------|--------|-----------|
| `tipos/detector.py` | 135 | NOVO - Detecta tipo vínculo |
| `tipos/clt.py` | 140 | Processador CLT |
| `tipos/facultativo.py` | 95 | Processador Facultativo |
| `parsers/vinculos.py` | 218 | Parse vínculos |
| `parsers/cabecalho.py` | 102 | NOVO - Parse cabeçalho |
| `orquestrador.py` | 175 | NOVO - Coordena extração |
| `utils.py` | 160 | Funções auxiliares |
| `__init__.py` | 78 | Entry points públicos |
| **TOTAL** | **1103** | |

### Comparação

| Versão | Arquivos | Linhas | Modularização |
|--------|----------|--------|---------------|
| Antes (v1.0) | 1 | 922 | ❌ Monolítico |
| Agora (v2.0) | 10 | 1103 | ✅ Modular por tipo |
| Ganho | +9 | +181 | +100% organização |

**Observação:** Aumento de linhas é esperado (docstrings, separação, imports). Benefício está na organização e manutenibilidade.

## 🎯 Próximos Passos (Ordem)

### Passo 1: Refatorar extrair_remuneracoes_texto() ⚠️ CRÍTICO
**Tempo estimado:** 1h  
**Prioridade:** ALTA (bloqueio para integração)

Criar `extrator/tipos/coordenador_remuneracoes.py` que:
1. Itera páginas do PDF
2. Detecta tipo de cada bloco de vínculo
3. Delega para processador específico (CLT, Facultativo)
4. Retorna lista consolidada de remunerações

### Passo 2: Integrar orquestrador no CLI
**Tempo estimado:** 30min  
**Prioridade:** ALTA

Modificar `converter_extrato_inss.py` para:
1. Importar `processar_cnis_completo`
2. Substituir chamadas inline por orquestrador
3. Simplificar main() para ~50 linhas

### Passo 3: Testar e validar
**Tempo estimado:** 15min  
**Prioridade:** CRÍTICA

1. Executar teste com CNIS_JOAO_CARLOS.pdf
2. Validar 178 remunerações
3. Comparar saídas com versão anterior

### Passo 4: Limpar e commitar
**Tempo estimado:** 10min  
**Prioridade:** MÉDIA

1. Remover arquivos legado (extrator_*.py)
2. Atualizar ARCHITECTURE.md
3. Commit final da Sessão 3

### Passo 5: Expandir tipos (futuro)
**Tempo estimado:** 2h cada  
**Prioridade:** BAIXA

Implementar novos tipos:
- [ ] `tipos/mei.py` - Microempreendedor Individual
- [ ] `tipos/autonomo.py` - Contribuinte com CNPJ próprio
- [ ] `tipos/segurado_especial.py` - Regime especial
- [ ] Atualizar `detector.py` com novos padrões

## 📝 Notas

### Decisões de Design

1. **Separação por Tipo:** Cada tipo de vínculo em arquivo próprio (vs classe base comum)
   - **Razão:** Prioridade é modularização clara. Reutilização vem depois.
   - **Trade-off:** Alguma duplicação de código aceitável nesta fase.

2. **Orquestrador Simples:** Coordena fluxo sem lógica complexa
   - **Razão:** Facilita entendimento e manutenção.
   - **Princípio:** Single Responsibility (cada função faz uma coisa).

3. **Detector Enum:** Tipo de vínculo como enum (vs strings)
   - **Razão:** Type safety, autocomplete, fácil adicionar tipos.
   - **Benefício:** `if tipo == TipoVinculo.CLT` vs `if tipo == "CLT"`.

### Lições Aprendidas

1. **Tentativa de Integração Direta Falhou:** 
   - Primeira tentativa de refatorar converter_extrato_inss.py quebrou extração (0 remunerações).
   - **Solução:** Criar estrutura modular separada primeiro, integrar depois com cuidado.

2. **Importância de Testes:**
   - Sem teste automatizado, difícil validar mudanças.
   - **Próximo:** Criar `pytest` para cada módulo.

3. **Documentação é Crucial:**
   - README_MODULAR.md ajuda entender filosofia e estrutura.
   - Facilita colaboração e manutenção futura.

## 🔗 Referências

- [README_MODULAR.md](README_MODULAR.md) - Arquitetura detalhada
- [ARCHITECTURE.md](../ARCHITECTURE.md) - Visão geral do sistema
- [PLANO_SIMPLIFICACAO.md](../PLANO_SIMPLIFICACAO.md) - Roadmap completo
- Commit 5c2f10d - Criação módulos base
- Commit 43c2739 - Reorganização modular por tipos
