════════════════════════════════════════════════════════════════
  📋 ROADMAP - PRÓXIMAS FASES DO PROJETO
════════════════════════════════════════════════════════════════

🎯 FASE ATUAL: PILOTO (v1.0.0)
─────────────────────────────
Status: ⏳ Aguardando validação advogado
Enviado: 27/01/2026 via WhatsApp
Escopo: Vínculos REMUNERAÇÃO + PRÉ-FACULTATIVO apenas


📅 FASE 2: AJUSTES PÓS-VALIDAÇÃO
────────────────────────────────
Prazo: Após retorno do advogado (1-2 semanas?)

Ações previstas:
[ ] Corrigir cálculos conforme orientação jurídica
[ ] Ajustar carência (15 vs 20 anos - confirmar)
[ ] Validar fator previdenciário (Lei 9.876/99)
[ ] Validar coeficiente pós-reforma (EC 103/2019)
[ ] Atualizar piso INSS para 2026 (R$ 1.412 → ?)
[ ] Revisar layout das abas se necessário
[ ] Incluir/remover dados conforme feedback

Critério de aprovação:
✓ Advogado confirma que cálculos estão corretos
✓ Comparação manual vs sistema bate 100%
✓ Nenhum ajuste jurídico pendente


📅 FASE 3: EXPANSÃO DE VÍNCULOS
───────────────────────────────
Prazo: 2-3 semanas após Fase 2

Objetivo: Suportar todos os tipos de vínculo do CNIS

Novos tipos a implementar:
[ ] Facultativo (contribuinte individual)
[ ] Empregado doméstico
[ ] Contribuinte individual (CI)
[ ] Segurado especial (rural)
[ ] Tempo militar
[ ] Tempo de estudante
[ ] RPPS (servidor público)
[ ] Períodos sem remuneração

Desafios técnicos:
• Cada tipo tem regras específicas de cálculo
• Conversão de tempo especial (fatores 1.2, 1.4)
• Averbação de tempo RPPS → RGPS
• Períodos com múltiplos vínculos simultâneos

Entrega:
→ Extrator universal que aceita qualquer tipo de CNIS
→ Documentação de regras por tipo de vínculo
→ Testes com casos reais diversos


📅 FASE 4: REGRAS DE TRANSIÇÃO
──────────────────────────────
Prazo: Paralelo à Fase 3

Objetivo: Implementar todas as regras de transição EC 103/2019

Regras a adicionar:
[ ] Pontos (86/96) - soma idade + tempo
[ ] Idade progressiva - tabela ano a ano
[ ] Pedágio 50% - quem estava próximo (2 anos)
[ ] Pedágio 100% - idade mínima + pedágio
[ ] Professor(a) - redução 5 anos idade

Complexidade:
⚠️ Alto - múltiplas regras simultâneas
⚠️ Requer lógica para escolher melhor opção
⚠️ Cliente pode se enquadrar em várias regras

Entrega:
→ Comparação automática de TODAS as opções
→ Recomendação da regra mais vantajosa
→ Tabela comparativa com datas e valores


📅 FASE 5: TEMPO ESPECIAL
─────────────────────────
Prazo: Após Fase 3

Objetivo: Converter tempo especial (insalubridade/periculosidade)

Conversões:
[ ] 15 anos especial × 1.67 = 25 anos comum
[ ] 20 anos especial × 1.40 = 28 anos comum
[ ] 25 anos especial × 1.20 = 30 anos comum

Necessário:
• Identificar períodos especiais no CNIS
• PPP (Perfil Profissiográfico Previdenciário)
• Laudo técnico de condições de trabalho
• Validação jurídica da conversão

Entrega:
→ Cálculo automático de tempo convertido
→ Simulações com e sem conversão
→ Documentação necessária para comprovar


📅 FASE 6: INTEGRAÇÃO VBA
─────────────────────────
Prazo: Após aprovação completa das Fases 2-5

Objetivo: Migrar sistema validado para VBA (erp_prev.xlsm)

Ações:
[ ] Portar lógica Python → VBA
[ ] Integrar com fluxo atual do escritório
[ ] Manter compatibilidade com processos existentes
[ ] Criar interface amigável para operadores
[ ] Documentar código VBA
[ ] Treinar equipe

Desafios:
• VBA é mais limitado que Python
• Performance pode ser menor
• Debugging mais difícil
• Manutenção futura

Entrega:
→ Sistema integrado ao erp_prev.xlsm
→ Manual de operação para equipe
→ Treinamento presencial/remoto


📅 FASE 7: FEATURES AVANÇADAS
─────────────────────────────
Prazo: Longo prazo (6+ meses)

Ideias futuras:
[ ] Simulador interativo (web?)
[ ] Alertas automáticos (cliente próximo de se aposentar)
[ ] Integração com Meu INSS (API oficial se existir)
[ ] OCR melhorado para PDFs de má qualidade
[ ] Geração automática de petições
[ ] Dashboard executivo para advogado
[ ] Relatórios em PDF formatados
[ ] Integração com sistema de gestão de clientes


🎯 MÉTRICAS DE SUCESSO
──────────────────────
• Redução de tempo de análise: 2-3 horas → 15 minutos
• Acurácia: 100% dos cálculos corretos
• Tipos de casos: Cobrir 95%+ dos casos reais
• Satisfação: Advogado aprova e usa regularmente
• Produtividade: 3-5× mais casos analisados/dia


📊 RECURSOS NECESSÁRIOS
───────────────────────
Fase 2: 1-2 dias (ajustes pontuais)
Fase 3: 2-3 semanas (desenvolvimento complexo)
Fase 4: 2-3 semanas (regras de transição)
Fase 5: 1-2 semanas (tempo especial)
Fase 6: 3-4 semanas (migração VBA + testes)
Fase 7: A definir (features futuras)

Total estimado: 3-4 meses para sistema completo


🚀 CRITÉRIO DE "PRONTO PARA PRODUÇÃO"
─────────────────────────────────────
✓ Advogado valida 100% dos cálculos
✓ Testes com 20+ casos reais diversos
✓ Todos os tipos de vínculo suportados
✓ Regras de transição implementadas
✓ Documentação completa
✓ Treinamento da equipe realizado
✓ Período de testes em paralelo (1-2 meses)
✓ Zero falhas críticas encontradas


═══════════════════════════════════════════════════════════════
Projeto: Sistema de Análise Previdenciária
Cliente: Escritório de Advocacia (25+ anos)
Data inicial: Janeiro 2026
Status: Fase 1 (Piloto) - Aguardando validação
═══════════════════════════════════════════════════════════════
