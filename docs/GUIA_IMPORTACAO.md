# Guia de Importação de Extratos do INSS

## Visão Geral

Este guia explica como preparar e importar dados de vínculos empregatícios do INSS para o Sistema de Planejamento Previdenciário.

## Obtendo o Extrato do INSS

### Via Portal Meu INSS

1. Acesse https://meu.inss.gov.br
2. Faça login com CPF e senha (ou conta gov.br)
3. Clique em "Extrato de Contribuição (CNIS)"
4. Baixe o extrato em PDF ou visualize online
5. Anote os dados dos vínculos manualmente ou use OCR

### Via Atendimento Presencial

1. Dirija-se a uma agência do INSS
2. Solicite o CNIS (Cadastro Nacional de Informações Sociais)
3. Receba o documento impresso ou em PDF

## Preparando o Arquivo CSV

### Formato Obrigatório

O arquivo deve ser um CSV (Comma-Separated Values) com ponto-e-vírgula como separador:

```
Empresa;CNPJ;Data Início;Data Fim;Tipo Vínculo;Condição Especial
```

### Estrutura das Colunas

| Coluna | Descrição | Formato | Obrigatório |
|--------|-----------|---------|-------------|
| Empresa | Nome da empresa/empregador | Texto | Sim |
| CNPJ | CNPJ do empregador | 00.000.000/0001-00 | Sim |
| Data Início | Início do vínculo | dd/mm/aaaa | Sim |
| Data Fim | Fim do vínculo | dd/mm/aaaa | Não* |
| Tipo Vínculo | Tipo de relação trabalhista | Texto | Sim |
| Condição Especial | Condição de trabalho | Texto | Sim |

\* Deixe vazio se o vínculo ainda está ativo

### Tipos de Vínculo Aceitos

- `CLT` - Consolidação das Leis do Trabalho
- `Estatutário` - Servidor público
- `Autônomo` - Contribuinte individual
- `MEI` - Microempreendedor Individual
- `Rural` - Trabalhador rural
- `Doméstico` - Empregado doméstico

### Condições Especiais

- `Normal` - Trabalho em condições normais
- `Insalubre` - Exposição a agentes insalubres
- `Periculoso` - Trabalho com periculosidade
- `Perigoso` - Atividade perigosa
- `Penoso` - Trabalho penoso

## Exemplo de Arquivo CSV Completo

```csv
Empresa;CNPJ;Data Início;Data Fim;Tipo Vínculo;Condição Especial
Empresa ABC Comércio Ltda;12.345.678/0001-90;01/01/2000;31/12/2005;CLT;Normal
Indústria XYZ Metalúrgica SA;98.765.432/0001-10;01/01/2006;31/12/2010;CLT;Insalubre
Comércio 123 Ltda;11.222.333/0001-44;01/01/2011;31/12/2015;CLT;Normal
Empresa Atual Serviços Ltda;55.666.777/0001-88;01/01/2016;;CLT;Normal
Trabalho Autônomo;00.000.000/0000-00;01/01/2018;;Autônomo;Normal
```

## Criando o Arquivo CSV

### Método 1: Excel

1. Abra o Microsoft Excel
2. Crie uma nova planilha
3. Digite os dados nas colunas conforme o formato
4. Vá em **Arquivo** → **Salvar Como**
5. Escolha o tipo **CSV (separado por vírgulas) (*.csv)**
6. Abra o arquivo em um editor de texto (Notepad)
7. Substitua todas as vírgulas por ponto-e-vírgula (;)
8. Salve o arquivo

### Método 2: Google Sheets

1. Acesse Google Sheets
2. Crie uma nova planilha
3. Digite os dados nas colunas
4. Vá em **Arquivo** → **Fazer download** → **Valores separados por vírgula (.csv)**
5. Abra o arquivo em um editor de texto
6. Substitua vírgulas por ponto-e-vírgula (;)
7. Salve o arquivo

### Método 3: Editor de Texto

1. Abra o Notepad ou outro editor de texto simples
2. Digite os dados linha por linha, separando os campos com ponto-e-vírgula
3. Salve com extensão .csv

**Exemplo:**
```
Empresa;CNPJ;Data Início;Data Fim;Tipo Vínculo;Condição Especial
Empresa ABC;12.345.678/0001-90;01/01/2000;31/12/2005;CLT;Normal
```

## Importando o Arquivo

### Via Macro VBA

```vba
ModImportacaoINSS.ImportarVinculosCSV _
    cpfCliente:="123.456.789-00", _
    caminhoArquivo:="C:\caminho\completo\para\arquivo.csv"
```

### Passo a Passo

1. Certifique-se de que o cliente já está cadastrado
2. Salve o arquivo CSV em um local conhecido
3. Pressione `Alt + F8` para abrir a lista de macros
4. Crie uma macro personalizada ou modifique `ExemploImportarVinculos`
5. Informe o CPF do cliente e o caminho completo do arquivo
6. Execute a macro
7. Aguarde a mensagem de confirmação

## Verificação Pós-Importação

Após importar, verifique:

1. **Planilha Vínculos**
   - Todos os vínculos foram importados?
   - As datas estão corretas?
   - Os CNPJs estão formatados corretamente?

2. **Planilha CalculoTempo**
   - O tempo total foi calculado?
   - O tempo especial está correto (se aplicável)?
   - Os valores em anos/meses/dias fazem sentido?

3. **Compare com o CNIS Original**
   - Número de vínculos corresponde?
   - Períodos batem com o extrato do INSS?
   - Há períodos faltantes?

## Tratamento de Erros Comuns

### Erro: "Cliente não encontrado"
**Causa:** CPF não cadastrado no sistema  
**Solução:** Cadastre o cliente antes de importar vínculos

### Erro: "Arquivo não encontrado"
**Causa:** Caminho do arquivo incorreto  
**Solução:** Verifique o caminho completo, use barras invertidas `\` no Windows

### Erro: "Formato de data inválido"
**Causa:** Data no formato errado  
**Solução:** Use o formato dd/mm/aaaa (exemplo: 01/01/2000)

### Dados Importados Incorretamente
**Causa:** Separador errado no CSV  
**Solução:** Certifique-se de usar ponto-e-vírgula (;) como separador

### Caracteres Estranhos
**Causa:** Encoding do arquivo incorreto  
**Solução:** Salve o CSV com encoding ANSI ou UTF-8

## Dicas Importantes

### ✓ Boas Práticas

1. **Sempre faça backup** antes de importações grandes
2. **Teste com um arquivo pequeno** primeiro
3. **Revise manualmente** os dados após importação
4. **Mantenha o arquivo CSV original** como referência
5. **Use nomes de arquivo descritivos** (ex: `vinculos_joao_silva.csv`)

### ✗ Evite

1. Não use acentos nos nomes das empresas (podem causar problemas)
2. Não deixe linhas vazias no meio do arquivo
3. Não use vírgulas dentro dos campos (use apenas ponto-e-vírgula como separador)
4. Não abra o arquivo CSV no Excel após salvar (pode alterar o formato)

## Importação em Lote

Para importar vínculos de múltiplos clientes:

```vba
Sub ImportarVariosCPFs()
    Dim cpfs() As String
    Dim arquivos() As String
    Dim i As Integer
    
    cpfs = Split("123.456.789-00,987.654.321-00,111.222.333-44", ",")
    arquivos = Split("C:\vinculos1.csv,C:\vinculos2.csv,C:\vinculos3.csv", ",")
    
    For i = LBound(cpfs) To UBound(cpfs)
        ModImportacaoINSS.ImportarVinculosCSV cpfs(i), arquivos(i)
    Next i
    
    MsgBox "Importação em lote concluída!", vbInformation
End Sub
```

## Exportando Dados do Sistema

Para gerar um CSV dos vínculos já cadastrados:

1. Acesse a planilha **Vínculos**
2. Selecione todos os dados (incluindo cabeçalho)
3. Copie para uma nova planilha
4. Salve como CSV

Isso permite criar backups ou transferir dados entre sistemas.

## Integração com PPP (Perfil Profissiográfico Previdenciário)

Para atividades especiais, além do CNIS, é importante ter:

- PPP do trabalhador
- LTCAT (Laudo Técnico de Condições Ambientais do Trabalho)
- Documentação complementar

Esses documentos servem para comprovar a condição especial informada no CSV.

## Suporte

Em caso de dúvidas sobre importação:
1. Consulte este guia
2. Verifique os exemplos em `templates/exemplo_vinculos.csv`
3. Teste com dados fictícios primeiro
4. Abra uma issue no GitHub se o problema persistir

---

**Importante:** Este sistema é uma ferramenta auxiliar. Sempre valide os dados importados com os documentos oficiais do INSS antes de usar em processos judiciais ou administrativos.
