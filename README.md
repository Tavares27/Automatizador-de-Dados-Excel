# Automatizador de Dados Excel

Este projeto é um script Python que automatiza a leitura, processamento e escrita de dados em uma planilha Excel. Ele utiliza a biblioteca `openpyxl` para manipular arquivos Excel. O caso de teste está sendo utilizado
com uma matriz 16x8 e 9 diferentes listas, mas esses valores podem ser alterados pelo usuário a depender da planilha.

## Funcionalidades

- Verifica se o arquivo Excel especificado existe.
- Carrega a planilha e desmescla todas as células mescladas.
- Inicializa uma matriz 16x8 e a preenche com os dados da planilha.
- Classifica os dados em diferentes listas (GS1 a GS9) com base em faixas de valores (thresholds).
- Escreve as listas classificadas de volta na planilha em colunas específicas.

## Requisitos

- Biblioteca `openpyxl`

Você pode instalar a biblioteca `openpyxl` usando o seguinte comando:

```bash
pip install openpyxl
