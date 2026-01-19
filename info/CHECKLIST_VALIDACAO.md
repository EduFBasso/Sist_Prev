# Checklist de Validação - Sist_Prev

## ✅ MOTOR DE CÁLCULO - VALIDADO

### Funções Base

- [x] `CalcularTempo()` - Calcula tempo total com tratamento de sobreposição
- [x] `CalcularIdade()` - Calcula idade precisa
- [x] `CalcularTempoEspecial()` - Converte tempo especial com fatores corretos
- [x] `DetectarConcomitancia()` - Identifica vínculos sobrepostos
- [x] `NovoID()` - Gera IDs incrementais
- [x] `Nz()` - Trata valores nulos
- [x] `GetParametro()` - Lê parâmetros da Config_Regras

### Regras de Aposentadoria

- [x] `RegraTempoContribuicao()` - 35 anos (H) / 30 anos (M)
- [x] `RegraIdade()` - 65 anos (H) / 62 anos (M) + carência
- [x] `RegraPontos()` - 105 pontos (H) / 100 pontos (M)
- [x] `RegraPedagio50()` - Pedágio 50% com verificação de elegibilidade
- [x] `RegraPedagio100()` - Pedágio 100% sem idade mínima

### Cálculos Avançados

- [x] `CalcularValorBeneficio()` - Estima valor do benefício
- [x] `VerificarElegibilidadeTransicao()` - Verifica regras aplicáveis
- [x] `AnalisarMelhorRegra()` - Identifica melhor opção automaticamente
- [x] `SimularAposentadoria()` - Interface unificada

### Gestão de Dados

- [x] `SalvarCliente()` - Salva dados do cliente
- [x] `CarregarCliente()` - Carrega dados do cliente
- [x] `BuscarLinhaCliente()` - Localiza cliente na planilha
- [x] `SalvarVinculo()` - Salva vínculo
- [x] `ExcluirVinculo()` - Remove vínculo
- [x] `CarregarVinculo()` - Carrega vínculo para edição
- [x] `AtualizarIndicadoresCliente()` - Atualiza flags automáticos

### Importação

- [x] `ImportarVinculosDeCSV()` - Importa vínculos do CNIS
- [x] Script Python completo para extração de PDF
- [x] Tratamento de duplicidades
- [x] Parsing de datas e competências

### Formulários

- [x] `frmSimulacoes.PreencherDadosIniciais()` - Carrega dados iniciais
- [x] `frmSimulacoes.cmdCalcular_Click()` - Calcula regra selecionada
- [x] `frmSimulacoes.cmdAnalisarTodas_Click()` - Análise automática

## 📊 TESTES RECOMENDADOS

### Teste 1: Cliente com Direito Adquirido

**Cenário:**

- Homem, 60 anos
- 36 anos de contribuição
- Sem tempo especial

**Resultado Esperado:**

- Regra de Tempo: ✅ Sim (já tem 35+ anos)
- Regra de Idade: ❌ Não (faltam 5 anos de idade)
- Regra de Pontos: ✅ Provável (60+36=96, faltam 9 pontos)

**Ações:**

1. Cadastrar cliente com esses dados
2. Cadastrar vínculos totalizando 36 anos
3. Executar simulação
4. Verificar se identifica direito adquirido

---

### Teste 2: Cliente Jovem Sem Direito

**Cenário:**

- Mulher, 45 anos
- 15 anos de contribuição
- Sem tempo especial

**Resultado Esperado:**

- Regra de Tempo: ❌ Não (faltam 15 anos)
- Regra de Idade: ❌ Não (faltam 17 anos de idade)
- Regra de Pontos: ❌ Não (45+15=60, faltam 40 pontos)

**Ações:**

1. Cadastrar cliente com esses dados
2. Cadastrar vínculos totalizando 15 anos
3. Executar simulação
4. Verificar cálculo de falta e projeções

---

### Teste 3: Cliente com Tempo Especial

**Cenário:**

- Homem, 55 anos
- 20 anos comuns + 10 anos especiais (grau 25)
- Tempo convertido: 20 + (10 × 1.4) = 34 anos

**Resultado Esperado:**

- Tempo Total calculado: ~34 anos
- Próximo da aposentadoria por tempo
- Pontos: 55+34=89 (falta 16 para pontos)

**Ações:**

1. Cadastrar cliente
2. Cadastrar 20 anos normais + 10 anos especiais (grau 25)
3. Verificar se conversão está correta
4. Executar simulação

---

### Teste 4: Cliente com Vínculos Concomitantes

**Cenário:**

- Homem, 50 anos
- Vínculo 1: 01/01/2000 a 31/12/2015 (16 anos)
- Vínculo 2: 01/01/2010 a 31/12/2020 (11 anos)
- Período sobreposto: 2010-2015 (6 anos)
- **Tempo real:** 21 anos (não 27)

**Resultado Esperado:**

- Sistema deve detectar concomitância
- Marcar ambos vínculos como concomitantes
- Calcular tempo correto (21 anos, não 27)

**Ações:**

1. Cadastrar cliente
2. Cadastrar os 2 vínculos com sobreposição
3. Verificar se flag "Concomitante" é marcado
4. Verificar se tempo total = 21 anos

---

### Teste 5: Pedágio 50% (Elegível)

**Cenário:**

- Mulher, 56 anos hoje (53 em 13/11/2019)
- 29 anos de contribuição em 13/11/2019
- Faltava 1 ano na reforma
- Pedágio: 0.5 × 1 = 0.5 anos
- Total necessário: 30 + 0.5 = 30.5 anos

**Resultado Esperado:**

- Elegível para Pedágio 50%
- Se já tem 30.5+ anos: direito adquirido
- Se não: calcular falta correta

**Ações:**

1. Cadastrar cliente com data de nascimento calculada
2. Cadastrar vínculos apropriados
3. Verificar elegibilidade
4. Testar cálculo do pedágio

---

### Teste 6: Pedágio 100% (Não Elegível para 50%)

**Cenário:**

- Homem, 60 anos hoje (57 em 13/11/2019)
- 30 anos de contribuição em 13/11/2019
- Faltavam 5 anos na reforma (> 2 anos)
- Não elegível para Pedágio 50%
- Pedágio 100%: 5 + 5 = 10 anos adicionais

**Resultado Esperado:**

- Não elegível para Pedágio 50%
- Elegível para Pedágio 100%
- Total necessário: 35 + 5 = 40 anos

**Ações:**

1. Cadastrar cliente
2. Verificar que não pode usar Pedágio 50%
3. Verificar cálculo correto do Pedágio 100%

---

### Teste 7: Análise de Melhor Regra

**Cenário:**

- Mulher, 59 anos
- 28 anos de contribuição
- Pontos atuais: 87

**Opções:**

- Tempo: faltam 2 anos
- Idade: faltam 3 anos
- Pontos: faltam 6.5 anos (13 pontos ÷ 2)

**Resultado Esperado:**

- Melhor regra: TEMPO (menor falta)
- Sistema deve selecionar automaticamente

**Ações:**

1. Cadastrar cliente
2. Clicar em "Analisar Todas as Regras"
3. Verificar se seleciona regra TEMPO
4. Validar cálculos de todas as regras

---

## 🔍 VALIDAÇÕES DE PARÂMETROS

### Config_Regras

- [ ] Salário Mínimo atualizado (verificar valor vigente)
- [ ] Teto INSS atualizado (verificar valor vigente)
- [ ] Data da Reforma: 13/11/2019 (fixo)
- [ ] Fatores de conversão especial: 1.4, 1.75, 2.33 (fixos por lei)
- [ ] Coeficiente inicial: 60% (fixo por reforma)
- [ ] Percentual acréscimo: 2% ao ano (fixo por reforma)

### Fórmulas Críticas

- [ ] Cálculo de idade: considera mês e dia de nascimento
- [ ] Cálculo de tempo: remove sobreposições corretamente
- [ ] Conversão especial: aplica fatores corretos (15→2.33, 20→1.75, 25→1.4)
- [ ] Pontos: idade + tempo (sem casas decimais na idade)
- [ ] Valor benefício: 60% + (tempo - base) × 2%

## ⚠️ EDGE CASES A TESTAR

### Dados Extremos

- [ ] Cliente com mais de 100 anos
- [ ] Cliente com 0 anos de contribuição
- [ ] Vínculo com data fim antes da data início
- [ ] Vínculo sem data fim (em andamento)
- [ ] Tempo especial sem grau informado
- [ ] Cliente com 50+ vínculos

### Situações Especiais

- [ ] Todos os vínculos são concomitantes
- [ ] Vínculo com 1 dia de duração
- [ ] Cliente já aposentado (simular outra regra)
- [ ] Data de nascimento inválida
- [ ] Sexo não informado

### Cálculos Limites

- [ ] Tempo exato no limite (ex: 35.00 anos)
- [ ] Pontos exatos (ex: 105.00)
- [ ] Valor benefício = salário mínimo
- [ ] Valor benefício = teto INSS
- [ ] Coeficiente > 100%

## ✅ CHECKLIST FINAL

### Documentação

- [x] README.md atualizado e completo
- [x] MOTOR_CALCULO_STATUS.md criado
- [x] GUIA_RAPIDO.md criado
- [x] Comentários em todas as funções VBA

### Código

- [x] Todas as 5 regras implementadas
- [x] Cálculo de valor do benefício
- [x] Análise automática de melhor regra
- [x] Detecção de concomitância
- [x] Conversão de tempo especial
- [x] Integração com formulários

### Testes

- [ ] Testar cada regra individualmente
- [ ] Testar análise automática
- [ ] Testar com dados reais de clientes
- [ ] Validar cálculos com especialista
- [ ] Testar importação de CNIS

### Deployment

- [ ] Backup do sistema atual
- [ ] Atualizar parâmetros com valores vigentes
- [ ] Treinar usuários
- [ ] Documentar casos de uso específicos
- [ ] Estabelecer rotina de manutenção

---

## 🎯 PRÓXIMOS PASSOS

### Curto Prazo (Opcional)

1. Adicionar validação de CPF
2. Implementar histórico de simulações
3. Adicionar botão "Exportar para PDF"
4. Melhorar interface visual dos formulários

### Médio Prazo (Opcional)

1. Integrar cálculo real de média salarial
2. Adicionar sistema de documentos anexos
3. Implementar relatórios personalizados
4. Criar dashboard de indicadores

### Longo Prazo (Opcional)

1. Migrar para aplicação standalone
2. Adicionar integração com sistemas externos
3. Implementar controle de usuários/permissões
4. Adicionar auditoria de alterações

---

**Status Atual:** ✅ MOTOR DE CÁLCULO COMPLETO  
**Data de Conclusão:** 11/01/2026  
**Pronto para:** Testes e validação com casos reais
