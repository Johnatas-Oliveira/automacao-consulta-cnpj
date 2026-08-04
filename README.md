# Automação para Consulta de CNPJ´S

Automação em Python para consulta e enriquecimento de dados cadastrais de
CNPJs utilizando APIs públicas (BrasilAPI, com fallback para ReceitaWS).

A partir de uma lista de CNPJs (planilha `.xlsx`, `.xls` ou `.csv`), a
automação consulta razão social, situação cadastral, endereço, contato,
QSA (sócios) e outros dados cadastrais, devolvendo uma planilha Excel
organizada.

## Como usar

```bash
pip install -r requirements.txt
```

1. Coloque a planilha de entrada (qualquer nome de arquivo) dentro da
   pasta `input/`
2. Rode informando o nome **exato** da coluna que contém os CNPJs:

```bash
python consulta_cnpj.py "CNPJ"
```

3. O resultado é salvo automaticamente em `output/`, com nome
   `resultado_cnpj_AAAA-MM-DD_HHMM.xlsx`

Se você cancelar a execução no meio (Ctrl+C), a planilha é gerada mesmo
assim com os resultados parciais obtidos até aquele momento (sufixo
`_PARCIAL`).

## Como funciona

1. Localiza a planilha dentro de `input/` (o nome do arquivo não importa)
2. Limpa, valida e remove CNPJs duplicados
3. Consulta cada CNPJ na [BrasilAPI](https://brasilapi.com.br/api/cnpj/v1)
   (fonte principal); em caso de falha ou CNPJ não encontrado, tenta o
   [ReceitaWS](https://receitaws.com.br) como fallback
4. Se a BrasilAPI encontrar a empresa mas vier sem email/telefone, tenta
   complementar esses dois campos consultando o ReceitaWS também
5. Faz até 3 tentativas por fonte em caso de timeout/erro de rede, com
   pausa entre consultas para não estourar o rate limit das APIs
   gratuitas

## Campos capturados

CNPJ, Razão Social, Nome Fantasia, Situação Cadastral, Data da Situação,
Natureza Jurídica, Porte, Capital Social, Data de Abertura, CNAE
Principal, Logradouro, Número, Bairro, Cidade, UF, CEP, Telefone, Email,
Sócios/QSA (nome e qualificação de cada sócio, separados por `;`) e
Status da consulta (Encontrado / Não encontrado / Erro)..
