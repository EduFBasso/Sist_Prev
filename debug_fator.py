"""Script para debug do cálculo do fator previdenciário"""

from analise_elegibilidade import calcular_idade, calcular_fator_previdenciario
import csv

# Ler dados do cliente
with open('saida/teste_apos_remocao_dados_cliente.csv', 'r', encoding='utf-8') as f:
    reader = csv.DictReader(f, delimiter=';')
    dados = next(reader)
    data_nasc = dados.get('DataNascimento', '')
    sexo = dados.get('Sexo', '').strip() or 'Masculino'

# Contar meses
with open('saida/teste_apos_remocao_remuneracoes.csv', 'r', encoding='utf-8') as f:
    reader = csv.DictReader(f, delimiter=';')
    tempo_meses = sum(1 for _ in reader)

# Calcular idade
idade_info = calcular_idade(data_nasc)
idade_anos = idade_info['anos'] + idade_info['meses'] / 12
tempo_anos = tempo_meses / 12

print('='*70)
print('DADOS PARA CÁLCULO DO FATOR PREVIDENCIÁRIO')
print('='*70)
print(f'Nome: {dados.get("Nome", "")}')
print(f'Data Nascimento: {data_nasc}')
print(f'Sexo: {sexo}')
print(f'')
print(f'Idade atual: {idade_info["anos"]} anos e {idade_info["meses"]} meses')
print(f'Idade decimal: {idade_anos:.4f} anos')
print(f'')
print(f'Tempo contribuição: {tempo_meses} meses')
print(f'Tempo decimal: {tempo_anos:.4f} anos')
print(f'')

# Calcular fator
fator = calcular_fator_previdenciario(tempo_anos, idade_anos, sexo)

print('='*70)
print('CÁLCULO DO FATOR PREVIDENCIÁRIO (LEI 9.876/99)')
print('='*70)
print(f'')
print(f'Fórmula: f = (Tc × a / Es) × [1 + (Id + Tc × a) / 100]')
print(f'')
print(f'Onde:')
print(f'  Tc = Tempo contribuição = {tempo_anos:.4f} anos')
print(f'  a  = Alíquota = 0.31')
print(f'  Id = Idade = {idade_anos:.4f} anos')
print(f'  Es = Expectativa sobrevida (tabela IBGE)')
print(f'')

# Calcular Es manualmente
idade_int = int(idade_anos)
if sexo == 'Masculino':
    tabela_es = {
        40: 38.5, 45: 33.9, 50: 29.5, 51: 28.6, 52: 27.7, 53: 26.8,
        54: 26.0, 55: 25.1, 56: 24.3, 57: 23.4, 58: 22.6, 59: 21.8,
        60: 21.0, 61: 20.2, 62: 19.4, 63: 18.7, 64: 17.9, 65: 17.2
    }
else:
    tabela_es = {
        40: 43.2, 45: 38.4, 50: 33.8, 51: 32.9, 52: 32.0, 53: 31.1,
        54: 30.2, 55: 29.3, 56: 28.5, 57: 27.6, 58: 26.7, 59: 25.9,
        60: 25.0, 61: 24.2, 62: 23.4, 63: 22.6, 64: 21.8, 65: 21.0
    }

Es = tabela_es.get(idade_int, 20.0)
print(f'  Es = {Es} anos (tabela IBGE para {sexo}, {idade_int} anos)')
print(f'')

# Calcular passo a passo
Tc = tempo_anos
a = 0.31
Id = idade_anos

parte1 = (Tc * a) / Es
parte2 = 1 + ((Id + (Tc * a)) / 100)
fator_calculado = parte1 * parte2

print('PASSO A PASSO:')
print(f'')
print(f'1) Tc × a = {Tc:.4f} × 0.31 = {Tc * a:.4f}')
print(f'')
print(f'2) Parte 1 = (Tc × a) / Es')
print(f'           = {Tc * a:.4f} / {Es}')
print(f'           = {parte1:.6f}')
print(f'')
print(f'3) Id + (Tc × a) = {Id:.4f} + {Tc * a:.4f} = {Id + Tc * a:.4f}')
print(f'')
print(f'4) Parte 2 = 1 + [(Id + Tc × a) / 100]')
print(f'           = 1 + [{Id + Tc * a:.4f} / 100]')
print(f'           = 1 + {(Id + Tc * a) / 100:.6f}')
print(f'           = {parte2:.6f}')
print(f'')
print(f'5) Fator = Parte 1 × Parte 2')
print(f'         = {parte1:.6f} × {parte2:.6f}')
print(f'         = {fator_calculado:.6f}')
print(f'')
print(f'6) Fator final (limitado): {fator:.4f}')
print(f'   (mínimo 0.4, máximo 1.3)')
print(f'')
print('='*70)
print('INTERPRETAÇÃO')
print('='*70)
if fator < 0.7:
    print('⚠️  FATOR BAIXO: Reduz significativamente o benefício')
    print('    → Aposentadoria precoce (pouca idade + pouco tempo)')
    print(f'    → Benefício = Média × {fator:.4f}')
elif fator < 1.0:
    print('⚠️  FATOR REDUTOR: Reduz o benefício')
    print('    → Idade/tempo insuficientes para fator neutro')
    print(f'    → Benefício = Média × {fator:.4f}')
elif fator == 1.0:
    print('✓  FATOR NEUTRO: Não altera o benefício')
    print('    → Equilíbrio entre idade e tempo de contribuição')
    print(f'    → Benefício = Média × {fator:.4f}')
else:
    print('✓  FATOR AMPLIADOR: Aumenta o benefício')
    print('    → Idade avançada + tempo longo de contribuição')
    print('    → Premia quem trabalhou mais tempo')
    print(f'    → Benefício = Média × {fator:.4f}')
print('='*70)

print()
print('OBSERVAÇÕES IMPORTANTES:')
print()
print('1) O fator previdenciário foi criado para desestimular')
print('   aposentadorias precoces (pouca idade + pouco tempo)')
print()
print('2) Com 52 anos e ~15 anos de contribuição:')
print(f'   - Idade ainda é considerada "jovem" para aposentadoria')
print(f'   - Tempo de contribuição é BAIXO (precisa 35 anos)')
print(f'   - Por isso o fator pode estar AMPLIANDO (acima de 1.0)')
print()
print('3) PROBLEMA: Essa análise assume que ele PODERIA se aposentar')
print('   hoje pela regra antiga, mas ELE NÃO PODE!')
print('   - Precisa completar 35 anos (420 meses)')
print('   - Faltam 242 meses!')
print()
print('4) O fator REAL seria calculado quando ele atingir 35 anos:')
print(f'   - Idade futura: ~72 anos')
print(f'   - Tempo: 35 anos')
print(f'   - Fator futuro: provavelmente MAIOR que 1.0')
print()
print('='*70)
