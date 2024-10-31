# Excel Data Sync

## Descrição
Este script automatiza o processamento e sincronização de dados entre as abas "Dados LI" e "Anuências" em uma planilha Excel. Ele desmescla células, remove linhas vazias e copia dados entre abas com base em IDs comuns (números de LI). O script é útil para garantir que dados relacionados estejam corretamente sincronizados em uma única planilha.

## Funcionalidades
- Remoção de células mescladas em ambas as abas.
- Eliminação de linhas com IDs de LI ausentes.
- Mapeamento de IDs entre abas e transferência de dados correspondente.
- Salvamento automático em uma nova planilha (`DadosFinal.xlsx`).

## Como usar
1. Coloque o arquivo Excel `Dados.xlsx` na pasta `planilhas`.
2. Execute o script.
3. O arquivo atualizado será salvo como `DadosFinal.xlsx` na mesma pasta.

## Dependências
- `openpyxl`

## Instalação
```base
Execute `pip install openpyxl`

