# Guia de Início Rápido

## 5 Minutos para Começar

### Passo 1: Preparar o Excel (1 min)

1. Abra o Microsoft Excel
2. Crie nova pasta de trabalho
3. Salve como `Sistema_Previdenciario.xlsm` (com macros)
4. Habilite a guia Desenvolvedor:
   - Arquivo → Opções → Personalizar Faixa de Opções
   - Marque "Desenvolvedor"

### Passo 2: Importar Módulos VBA (2 min)

1. Pressione `Alt + F11` (abre Editor VBA)
2. Menu Arquivo → Importar Arquivo
3. Importe os 4 arquivos de `src/modules/`:
   - ModPrincipal.bas
   - ModClientes.bas
   - ModImportacaoINSS.bas
   - ModSimulacaoAposentadoria.bas

### Passo 3: Inicializar Sistema (30 seg)

1. Pressione `Alt + F8` (lista de macros)
2. Selecione `InicializarSistema`
3. Clique em Executar
4. Aguarde a mensagem de sucesso

### Passo 4: Executar Demonstração (1 min)

1. Pressione `Alt + F8`
2. Execute `DemonstracaoCompleta`
3. Veja os dados de exemplo criados
4. Navegue pelas planilhas criadas

### Passo 5: Explorar o Sistema (30 seg)

Veja as planilhas criadas:
- **Início** - Dashboard com instruções
- **Clientes** - Cliente de exemplo cadastrado
- **Vínculos** - Vínculo de exemplo adicionado
- **CalculoTempo** - Cálculo automático realizado
- **ResultadosSimulacao** - Simulação de aposentadorias

## Usar com Seus Dados

### Cadastrar Seu Primeiro Cliente Real

```vba
' Pressione Alt + F11, crie um novo módulo, cole e execute:

Sub MeuPrimeiroCliente()
    ModClientes.CadastrarCliente _
        nome:="[NOME COMPLETO]", _
        cpf:="[000.000.000-00]", _
        dataNascimento:=DateSerial(1970, 1, 15), _
        sexo:="M", _
        email:="[email@exemplo.com]", _
        telefone:="([00]) [00000-0000]"
End Sub
```

Substitua os valores entre colchetes e execute!

### Adicionar Vínculos

```vba
Sub MeuPrimeiroVinculo()
    ModImportacaoINSS.AdicionarVinculo _
        cpfCliente:="[CPF DO CLIENTE]", _
        empresa:="[Nome da Empresa]", _
        cnpj:="[00.000.000/0001-00]", _
        dataInicio:=DateSerial(2000, 1, 1), _
        dataFim:=DateSerial(2020, 12, 31), _
        tipoVinculo:="CLT", _
        condicaoEspecial:="Normal"
    
    ' Calcular tempo após adicionar vínculo
    ModImportacaoINSS.CalcularTemposContribuicao "[CPF DO CLIENTE]"
End Sub
```

### Simular Aposentadoria

```vba
Sub MinhaSimulacao()
    Dim resultado As String
    resultado = ModSimulacaoAposentadoria.SimularAposentadorias("[CPF DO CLIENTE]")
    MsgBox resultado, vbInformation, "Resultado"
End Sub
```

## Importar CSV

1. Prepare um arquivo CSV usando o template em `templates/exemplo_vinculos.csv`
2. Use este código:

```vba
Sub ImportarMeuCSV()
    ModImportacaoINSS.ImportarVinculosCSV _
        cpfCliente:="[CPF DO CLIENTE]", _
        caminhoArquivo:="C:\caminho\para\arquivo.csv"
End Sub
```

## Próximos Passos

- 📖 Leia o [Manual do Usuário](MANUAL_USUARIO.md) completo
- 📊 Consulte o [Guia de Importação](GUIA_IMPORTACAO.md) para trabalhar com extratos do INSS
- 💡 Explore as planilhas e personalize conforme necessário

## Atalhos Úteis

| Atalho | Função |
|--------|--------|
| `Alt + F8` | Listar macros |
| `Alt + F11` | Abrir Editor VBA |
| `Ctrl + PageDown/Up` | Navegar entre planilhas |
| `F5` | Executar macro no VBA |

## Dúvidas Frequentes

**P: As macros não funcionam**  
R: Verifique se salvou como .xlsm e habilitou macros

**P: Erro "Planilha não encontrada"**  
R: Execute `InicializarSistema` primeiro

**P: Como apagar dados de exemplo?**  
R: Delete as linhas nas planilhas ou execute `InicializarSistema` em arquivo novo

**P: Posso usar no Mac?**  
R: Sim, mas o VBA tem algumas limitações no Excel para Mac

## Suporte

- 📘 Documentação completa: `/docs`
- 🐛 Reportar problemas: GitHub Issues
- 💬 Dúvidas: Consulte o README.md

---

**Pronto para começar!** Em 5 minutos você terá um sistema completo de planejamento previdenciário funcionando.
