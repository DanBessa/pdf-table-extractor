# PDF Table Extractor

Script para extrair dados de PDF do extrato consolidado do banco Itaú, o que facilita muito analistas contabeis e ou contadores no dia-a-dia


## Funcionalidades

- Extrai tabelas de múltiplas páginas de um PDF especificado.
- Processa páginas com diferentes configurações de área de tabela e colunas.
- Concatena dados de várias páginas em um único arquivo CSV.
- Ajusta e sanitiza os nomes das colunas.
- Remove colunas vazias ou com nomes indesejados.
- Alinha e combina CSVs gerados, removendo delimitadores consecutivos.

## Requisitos

- Python 3.6 ou superior
- Bibliotecas Python:
  - `camelot-py`
  - `pandas`
  - `unidecode`
  - `tkinter`

Para instalar as dependências, execute:

```bash
pip install camelot-py[pdf] pandas unidecode
```

## Como usar
1. Execute o script.

```bash
python pdf_table_extractor.py
```

2.Uma janela será aberta para você selecionar o arquivo PDF.

3.Em seguida, insira as páginas que contêm as tabelas. Use o formato:

 - Exemplo: 1,2,4-6 (extrai as páginas 1, 2 e da 4 à 6).
4.O script processará o PDF e salvará o CSV extraído no mesmo diretório que o arquivo PDF.

## Configurações
As configurações de extração, como áreas de tabela e colunas, podem ser ajustadas dentro do script na variável configs:

  - flavor: Define o modo de extração de tabela (ex: stream).
  - table_areas: Coordenadas para delimitar a área onde as tabelas estão localizadas em cada página.
  - columns: Coordenadas das colunas da tabela para a extração correta.
## Estrutura do Projeto
```bash
pdf_table_extractor/
│
├── pdf_table_extractor.py  # Código principal de extração
├── README.md               # Este arquivo
└── requirements.txt        # Dependências do projeto
```
## Contribuições
Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou enviar pull requests com melhorias.
