# 📋 PLANO DE ANÁLISE - PRÉ E PÓS-REFORMA 2019

## 🎯 OBJETIVO
Criar abas na planilha Excel que analisem automaticamente a elegibilidade do cliente para aposentadoria em **duas regras distintas**:
- **Regra Antiga** (antes EC 103/2019)
- **Regra Nova** (após EC 103/2019)

---

## 📊 ESTRUTURA DAS ABAS

### **ABA 7: Analise_Pre_Reforma**
Aposentadoria por tempo de contribuição (regra antiga - até 13/11/2019)

#### Seção 1: Dados Base
- Nome do cliente
- Data de nascimento
- Idade atual
- Sexo
- Data de referência (hoje)

#### Seção 2: Requisitos da Regra Antiga
**Aposentadoria por Tempo de Contribuição (Integral)**
- ❌ **Extinta pela Reforma** - Mantida apenas para direito adquirido
- Requisitos históricos (antes da reforma):
  - Homens: 35 anos de contribuição
  - Mulheres: 30 anos de contribuição
  - **Sem idade mínima** (diferença principal)

**Fator Previdenciário:**
- Fórmula: f = (Tc × a × [1 + (Id + Tc × a)] / Es
- Tc = Tempo de contribuição
- a = Alíquota (0,31)
- Id = Idade no momento da aposentadoria
- Es = Expectativa de sobrevida (tabela IBGE)

**Fórmula do Benefício:**
- Média = 80% maiores salários (desde 07/1994)
- Benefício = Média × Fator Previdenciário

#### Seção 3: Análise Atual
```
PODE SE APOSENTAR PELA REGRA ANTIGA?
[SIM] ou [NÃO]

SE SIM:
  - Tempo de contribuição: XXX anos (YYY meses)
  - Idade atual: ZZ anos
  - Fator previdenciário: 0.XXXX
  - Benefício estimado: R$ X.XXX,XX
  
SE NÃO:
  - Tempo atual: XXX anos (YYY meses)
  - Tempo necessário: 35 anos (homem) / 30 anos (mulher)
  - FALTA: ZZ meses
  
  SIMULAÇÃO CONTRIBUTIVA:
  - Contribuindo com salário mínimo (R$ X.XXX,XX):
    → Data prevista elegibilidade: DD/MM/AAAA
    → Benefício estimado: R$ X.XXX,XX
  
  - Contribuindo com média atual (R$ X.XXX,XX):
    → Data prevista elegibilidade: DD/MM/AAAA
    → Benefício estimado: R$ X.XXX,XX
  
  - Contribuindo para NÃO REDUZIR média (R$ X.XXX,XX):
    → Valor mensal necessário: R$ X.XXX,XX
    → Data prevista: DD/MM/AAAA
    → Benefício estimado: R$ X.XXX,XX
```

#### Seção 4: Regras de Transição (para quem estava perto)
**Regra 1: Pedágio de 50%**
- Requisitos:
  - Faltavam no máximo 2 anos para 35H/30M em 13/11/2019
  - Cumprir 50% do tempo faltante como "pedágio"
- Análise: ELEGÍVEL? [SIM/NÃO]
- Se NÃO: Faltavam X anos em 2019 (não se enquadra)

**Regra 2: Pedágio de 100%**
- Requisitos:
  - Idade mínima: 60H / 57M
  - Tempo mínimo: 35H / 30M
  - Pedágio: 100% do tempo faltante em 2019
- Análise: ELEGÍVEL? [SIM/NÃO]

**Regra 3: Idade Progressiva**
- Idade mínima: 61H / 56M (em 2019) → Aumenta 6 meses/ano
- Tempo: 35H / 30M
- Pontos: 96 (2019) → Aumenta 1 ponto/ano até 105
- Análise: ELEGÍVEL? [SIM/NÃO]

**Regra 4: Idade Mínima Progressiva**
- Idade: 61H/56M (2019) + 6 meses/ano até 65H/62M
- Tempo: 35H/30M
- Análise: ELEGÍVEL? [SIM/NÃO]

---

### **ABA 8: Analise_Pos_Reforma**
Aposentadoria programada (regra nova - após 13/11/2019)

#### Seção 1: Requisitos da Regra Nova
**Aposentadoria Programada (EC 103/2019)**
- Idade mínima: 65 anos (homens) / 62 anos (mulheres)
- Tempo mínimo: 20 anos de contribuição (ambos)
- Tempo para 100% da média: 35H / 30M

**Fórmula do Benefício:**
- Média = 100% de TODOS os salários (desde 07/1994)
- Coeficiente base: 60%
- Acréscimo: +2% por ano acima de 20H / 15M
- Coeficiente máximo: 100%
- Benefício = Média × Coeficiente

#### Seção 2: Análise Atual
```
PODE SE APOSENTAR PELA REGRA NOVA?
[SIM] ou [NÃO]

REQUISITOS:
  ✓ Idade mínima: [OK] ou [FALTA: X anos Y meses]
  ✓ Tempo mínimo: [OK] ou [FALTA: X anos Y meses]

SE SIM:
  - Idade: XX anos
  - Tempo de contribuição: YYY anos (ZZZ meses)
  - Coeficiente: XX% (60% + 2% × anos acima de 20)
  - Benefício: R$ X.XXX,XX
  
SE NÃO:
  CENÁRIO 1: Atingir idade mínima
    - Idade atual: XX anos
    - Falta: YY anos ZZ meses
    - Data prevista: DD/MM/AAAA
    - Coeficiente na data: XX%
    - Benefício estimado: R$ X.XXX,XX
  
  CENÁRIO 2: Atingir 100% da média (35H/30M)
    - Tempo atual: XX anos (YYY meses)
    - Falta: ZZ meses
    - Coeficiente: 100%
    - Benefício estimado: R$ X.XXX,XX
    
  CONTRIBUIÇÃO FACULTATIVA:
    - Valor mensal para NÃO reduzir média: R$ X.XXX,XX
    - Prazo até elegibilidade: ZZ meses
    - Benefício final estimado: R$ X.XXX,XX
```

#### Seção 3: Comparação de Cenários
```
MELHOR ESTRATÉGIA:

Opção A: Esperar idade mínima (sem contribuir)
  → Data: DD/MM/AAAA
  → Benefício: R$ X.XXX,XX
  → Custo: R$ 0,00

Opção B: Contribuir até 35H/30M (100% da média)
  → Data: DD/MM/AAAA
  → Benefício: R$ X.XXX,XX
  → Custo total: R$ X.XXX,XX
  → Ganho mensal: +R$ XXX,XX

Opção C: Contribuir com salário mínimo até idade
  → Data: DD/MM/AAAA
  → Benefício: R$ X.XXX,XX
  → Custo total: R$ X.XXX,XX
  → Ganho mensal: +R$ XXX,XX

RECOMENDAÇÃO: [Opção X] - Melhor custo-benefício
```

---

### **ABA 9: Comparacao_Geral**
Quadro comparativo final

```
═══════════════════════════════════════════════════════════════
                COMPARAÇÃO: PRÉ-REFORMA vs PÓS-REFORMA
═══════════════════════════════════════════════════════════════

┌─────────────────────────┬─────────────────┬─────────────────┐
│                         │   PRÉ-REFORMA   │   PÓS-REFORMA   │
│                         │   (até 2019)    │   (após 2019)   │
├─────────────────────────┼─────────────────┼─────────────────┤
│ Pode aposentar hoje?    │      [X]        │      [ ]        │
├─────────────────────────┼─────────────────┼─────────────────┤
│ Requisitos atendidos    │                 │                 │
│   - Idade               │  XX anos (OK)   │  XX (falta YY)  │
│   - Tempo contribuição  │  XX anos (OK)   │  XX anos (OK)   │
│   - Pontos              │  N/A            │  N/A            │
├─────────────────────────┼─────────────────┼─────────────────┤
│ Benefício calculado     │  R$ X.XXX,XX    │  R$ X.XXX,XX    │
│   - Base de cálculo     │  80% maiores    │  100% todos     │
│   - Fator/Coeficiente   │  0,XXXX         │  XX%            │
├─────────────────────────┼─────────────────┼─────────────────┤
│ Se NÃO elegível hoje    │                 │                 │
│   - Falta (meses)       │  XX meses       │  YY meses       │
│   - Data prevista       │  DD/MM/AAAA     │  DD/MM/AAAA     │
│   - Benefício futuro    │  R$ X.XXX,XX    │  R$ X.XXX,XX    │
└─────────────────────────┴─────────────────┴─────────────────┘

═══════════════════════════════════════════════════════════════
                          MELHOR OPÇÃO
═══════════════════════════════════════════════════════════════

➤ RECOMENDAÇÃO: [PRÉ-REFORMA] ou [PÓS-REFORMA]

MOTIVO:
  • Maior benefício: R$ X.XXX,XX vs R$ Y.YYY,YY
  • Menor tempo de espera: XX meses vs YY meses
  • Melhor custo-benefício: ...

AÇÕES NECESSÁRIAS:
  ☐ Requerer aposentadoria imediatamente
  ☐ Contribuir mais XX meses com R$ Y.YYY,YY/mês
  ☐ Aguardar até DD/MM/AAAA para atingir idade mínima
  ☐ Consultar advogado previdenciário para estratégia

═══════════════════════════════════════════════════════════════
```

---

## 🔧 IMPLEMENTAÇÃO TÉCNICA

### Etapa 1: Criar funções auxiliares
```python
def calcular_idade(data_nascimento, data_referencia=None):
    """Calcula idade em anos, meses e dias."""
    pass

def calcular_tempo_contribuicao(remuneracoes_df):
    """Calcula tempo total de contribuição."""
    pass

def calcular_fator_previdenciario(tempo_anos, idade_anos, sexo):
    """Calcula fator previdenciário (regra antiga)."""
    pass

def calcular_media_80_maiores(remuneracoes_df):
    """Calcula média dos 80% maiores salários."""
    pass

def verificar_elegibilidade_pre_reforma(dados_cliente, tempo_contrib):
    """Verifica se pode aposentar pela regra antiga."""
    pass

def verificar_elegibilidade_pos_reforma(dados_cliente, tempo_contrib):
    """Verifica se pode aposentar pela regra nova."""
    pass

def simular_contribuicao_facultativa(dados_atual, valor_mensal, meses):
    """Simula impacto de contribuições futuras."""
    pass
```

### Etapa 2: Criar abas na planilha
- `criar_aba_analise_pre_reforma(wb, dados_cliente, remuneracoes)`
- `criar_aba_analise_pos_reforma(wb, dados_cliente, remuneracoes)`
- `criar_aba_comparacao_geral(wb, analise_pre, analise_pos)`

### Etapa 3: Integrar com interface
- Adicionar botão "Análise Completa" (opcional)
- Ou sempre gerar automaticamente após criar planilha base

---

## 📅 CRONOGRAMA

**FASE 3.2 - Regra Pré-Reforma** (4-6 horas)
- Implementar cálculo fator previdenciário
- Implementar média 80% maiores
- Criar aba Analise_Pre_Reforma
- Testar com dados João Carlos

**FASE 3.3 - Regra Pós-Reforma** (3-4 horas)
- Implementar cálculo coeficiente
- Verificar elegibilidade idade + tempo
- Criar aba Analise_Pos_Reforma
- Testar com dados João Carlos

**FASE 3.4 - Comparação e Simulações** (2-3 horas)
- Implementar simulador contribuição facultativa
- Criar aba Comparacao_Geral
- Gerar recomendações automáticas

**FASE 3.5 - Validação Jurídica** (1-2 horas)
- Revisar com advogado
- Ajustar regras conforme legislação
- Documentar referências legais

---

## ⚖️ REFERÊNCIAS LEGAIS

- **EC 103/2019**: Reforma da Previdência (13/11/2019)
- **Lei 8.213/91**: Planos de Benefícios da Previdência Social
- **Lei 9.876/99**: Fator previdenciário
- **Decreto 3.048/99**: Regulamento da Previdência Social
- **IN INSS/PRES nº 128/2022**: Reconhecimento de direitos

---

## 💡 OBSERVAÇÕES IMPORTANTES

1. **Direito Adquirido**: Quem completou requisitos antes da reforma mantém direito à regra antiga
2. **Regras de Transição**: 5 modalidades diferentes para quem estava próximo em 2019
3. **Cálculo de Tempo**: Conversão de tempo especial, tempo rural, etc. (implementar depois)
4. **Revisões**: Cálculos podem ser revisados judicialmente (INPC vs SELIC)
5. **Teto/Piso**: Sempre aplicar limitadores (salário mínimo e teto INSS)

---

## ✅ PRÓXIMOS PASSOS

1. ✅ Ajustar botão "Concluído - Fechar" (FEITO)
2. ⏳ Validar este plano com você
3. ⏳ Implementar FASE 3.2 (Análise Pré-Reforma)
4. ⏳ Implementar FASE 3.3 (Análise Pós-Reforma)
5. ⏳ Implementar FASE 3.4 (Comparação)
6. ⏳ Validar com advogado

**Podemos começar?**
