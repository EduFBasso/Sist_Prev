# Testes Automatizados

Suite de testes para o conversor CNIS.

## Executar Testes

```bash
# Todos os testes
pytest tests/ -v

# Com cobertura
pytest tests/ --cov=converter_extrato_inss --cov=extrator --cov-report=html

# Teste específico
pytest tests/test_extrator.py::test_extracao_total_178_remuneracoes -v
```

## Resultados Atuais

✅ **12 testes passando**
- 6 testes de extração de remunerações
- 3 testes de geração de CSV
- 2 testes de casos extremos
- 1 teste de pipeline completo

📊 **Cobertura: 48%**
- converter_extrato_inss.py: 71%
- coordenador_remuneracoes.py: 88%
- config_vinculos.py: 88%

## Validações

### Extração
- ✅ Total de 178 remunerações (100%)
- ✅ Distribuição por Seq (1-13)
- ✅ Vínculos CLT com CNPJ válido
- ✅ 15 vínculos Facultativos (Seq 11-13)
- ✅ Formato de competência (MM/AAAA)
- ✅ Valores numéricos positivos

### Geração de CSV
- ✅ CSV de remunerações (178 linhas)
- ✅ CSV de dados do cliente (1 linha)
- ✅ CSV de vínculos estruturados (10 vínculos CLT)

### Casos Extremos
- ✅ PDF inexistente lança FileNotFoundError
- ✅ Ordenação por Seq

### Integração
- ✅ Pipeline completo (5 arquivos CSV gerados)
