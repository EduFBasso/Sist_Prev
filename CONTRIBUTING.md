# Guia de Contribuição

Obrigado por considerar contribuir para o Sistema de Planejamento Previdenciário RGPS!

## Como Contribuir

### Reportando Bugs

Se você encontrar um bug, por favor abra uma issue incluindo:

1. **Descrição clara do problema**
2. **Passos para reproduzir**
3. **Comportamento esperado vs comportamento atual**
4. **Versão do Excel e sistema operacional**
5. **Screenshots, se aplicável**

### Sugerindo Melhorias

Para sugerir novas funcionalidades:

1. Verifique se a sugestão já não existe nas issues
2. Descreva claramente o problema que a feature resolve
3. Explique como você imagina que funcionaria
4. Se possível, forneça exemplos de uso

### Pull Requests

1. **Fork o repositório**
2. **Crie uma branch para sua feature**
   ```bash
   git checkout -b feature/MinhaNovaFeature
   ```
3. **Faça suas alterações**
4. **Teste completamente suas mudanças**
5. **Commit suas alterações**
   ```bash
   git commit -m "feat: adiciona nova funcionalidade X"
   ```
6. **Push para sua branch**
   ```bash
   git push origin feature/MinhaNovaFeature
   ```
7. **Abra um Pull Request**

## Padrões de Código

### VBA

- Use `Option Explicit` em todos os módulos
- Comente código complexo em português
- Use nomes descritivos para variáveis
- Prefira funções a subrotinas quando retornar valor
- Sempre trate erros com `On Error GoTo`

**Exemplo:**
```vba
Option Explicit

' Função para calcular idade a partir da data de nascimento
Public Function CalcularIdade(dataNascimento As Date) As Integer
    On Error GoTo ErroHandler
    
    Dim idadeCalculada As Integer
    idadeCalculada = Year(Date) - Year(dataNascimento)
    
    ' Ajustar se ainda não fez aniversário no ano corrente
    If Month(Date) < Month(dataNascimento) Or _
       (Month(Date) = Month(dataNascimento) And Day(Date) < Day(dataNascimento)) Then
        idadeCalculada = idadeCalculada - 1
    End If
    
    CalcularIdade = idadeCalculada
    Exit Function
    
ErroHandler:
    MsgBox "Erro ao calcular idade: " & Err.Description
    CalcularIdade = 0
End Function
```

### Nomenclatura

- **Módulos:** `ModNomeDoModulo.bas`
- **Funções Públicas:** PascalCase (`CadastrarCliente`)
- **Funções Privadas:** PascalCase com prefixo (`PrivateCalcularTotal`)
- **Variáveis:** camelCase (`dataNascimento`, `cpfCliente`)
- **Constantes:** SNAKE_CASE_UPPER (`MAX_IDADE`, `TEMPO_MINIMO`)

### Estrutura de Módulos

```vba
Attribute VB_Name = "ModNome"
Option Explicit

' ========================================
' CONSTANTES
' ========================================

' ========================================
' FUNÇÕES PÚBLICAS
' ========================================

' ========================================
' FUNÇÕES PRIVADAS/AUXILIARES
' ========================================

' ========================================
' INICIALIZAÇÃO
' ========================================
```

## Documentação

### Comentários no Código

- Comente o **porquê**, não o **o quê**
- Use comentários de linha única para explicações breves
- Use blocos de comentários para seções

### Documentação Markdown

- Use português claro e objetivo
- Inclua exemplos práticos
- Formate código com syntax highlighting
- Use tabelas para comparações
- Adicione emojis para melhor visualização (✓, ✗, 📝, etc.)

## Testes

Antes de submeter um PR, teste:

1. **Inicialização do sistema** - Execute `InicializarSistema`
2. **Cadastro de cliente** - Teste com CPF válido e inválido
3. **Importação de vínculos** - CSV e manual
4. **Cálculo de tempos** - Verifique precisão
5. **Simulações** - Teste todos os tipos de aposentadoria
6. **Casos extremos:**
   - Datas futuras
   - CPF inválidos
   - Arquivos CSV malformados
   - Clientes sem vínculos

## Tipos de Commits

Use prefixos no estilo Conventional Commits:

- `feat:` Nova funcionalidade
- `fix:` Correção de bug
- `docs:` Alterações em documentação
- `style:` Formatação, ponto-e-vírgula faltando, etc.
- `refactor:` Refatoração de código
- `test:` Adição de testes
- `chore:` Atualização de tarefas, configurações, etc.

**Exemplos:**
```
feat: adiciona suporte para conversão de tempo especial
fix: corrige cálculo de pontos para aposentadoria
docs: atualiza manual do usuário com novos exemplos
refactor: simplifica validação de CPF
```

## Áreas que Precisam de Contribuição

### 🔥 Alta Prioridade

- [ ] Converter tempo especial em tempo comum
- [ ] Suporte para múltiplos vínculos simultâneos
- [ ] Interface gráfica (UserForms)
- [ ] Exportação de relatórios em PDF
- [ ] Validação cruzada com CNIS real

### 📊 Melhorias

- [ ] Gráficos de evolução de contribuições
- [ ] Projeções futuras (quando poderá se aposentar)
- [ ] Cálculo de benefício estimado
- [ ] Comparação entre diferentes modalidades
- [ ] Histórico de alterações nos dados

### 🎨 UX/UI

- [ ] Formulários visuais para cadastro
- [ ] Dashboard interativo
- [ ] Visualização de linha do tempo
- [ ] Temas/cores personalizáveis
- [ ] Ícones e formatação melhorada

### 📝 Documentação

- [ ] Vídeos tutoriais
- [ ] Casos de uso reais (anonimizados)
- [ ] FAQ expandido
- [ ] Troubleshooting detalhado
- [ ] Exemplos de integração com outros sistemas

## Roadmap

### Versão 1.1 (Próxima)
- Conversão de tempo especial
- Interface gráfica básica
- Exportação de relatórios

### Versão 1.2
- Gráficos e visualizações
- Projeções futuras
- Múltiplos vínculos simultâneos

### Versão 2.0
- Integração com APIs do INSS (se disponível)
- Sistema multi-usuário
- Banco de dados externo
- Versão web

## Código de Conduta

- Seja respeitoso e profissional
- Aceite críticas construtivas
- Foque no que é melhor para o projeto
- Mantenha discussões técnicas e objetivas

## Licença

Ao contribuir, você concorda que suas contribuições serão licenciadas sob a mesma licença MIT do projeto.

## Perguntas?

Se tiver dúvidas sobre como contribuir:

1. Leia a documentação em `/docs`
2. Procure em issues fechadas
3. Abra uma issue com a tag `question`
4. Entre em contato com os mantenedores

---

**Obrigado por contribuir para tornar o planejamento previdenciário mais acessível!** 🙏
