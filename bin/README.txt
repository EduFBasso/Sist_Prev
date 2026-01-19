Este diretório contém os executáveis standalone para distribuição ao cliente.

Os arquivos .exe são gerados pelo PyInstaller a partir dos scripts Python.

ARQUIVOS ESPERADOS:
- atualizar_inpc.exe      (~8 MB)  - Atualização de índices INPC do BCB
- converter_extrato_inss.exe (~15 MB) - Conversão de PDF CNIS para CSV

COMO GERAR:
Execute o script build_executaveis.py na pasta raiz do projeto.

IMPORTANTE:
- Cliente NÃO precisa ter Python instalado
- Executáveis funcionam standalone no Windows
- Antivírus pode alertar (falso positivo) - adicionar exceção se necessário
