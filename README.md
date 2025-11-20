# 📘 Processamento e Normalização de Endereços -- `script.py`

Este projeto realiza **normalização, deduplicação, extração de
coordenadas geográficas e parsing estruturado** de endereços contidos em
uma planilha Excel.\
O arquivo principal é **`script.py`**, que lê uma planilha chamada
`teste_Paulo.xlsx`, processa os endereços e gera um novo arquivo
tratado.

------------------------------------------------------------------------

## 🚀 Instruções para executar o projeto

### 1️⃣ Ative o ambiente virtual

**Linux / macOS**

``` bash
source venv/bin/activate
```

**Windows**

``` bash
venv\Scripts\activate
```

------------------------------------------------------------------------

### 2️⃣ Coloque o arquivo de entrada na mesma pasta

Certifique-se de que o arquivo:

    teste_Paulo.xlsx

está localizado no mesmo diretório onde está o arquivo:

    script.py

------------------------------------------------------------------------

### 3️⃣ Execute o script

Use um dos comandos abaixo:

``` bash
python3 script.py
```

ou

``` bash
py script.py
```

------------------------------------------------------------------------

## 🎯 Propósito do Projeto

Este script foi criado para **padronizar e estruturar endereços
mexicanos** a partir de uma fonte Excel para facilitar análises,
geocodificação e organização de dados.

As principais funcionalidades incluem:

-   Normalização de logradouros (ex: transformar "Av." → "AVENIDA").
-   Remoção de acentos e duplicações de espaços.
-   Extração automática de latitude e longitude presentes entre
    parênteses.
-   Deduplicação inteligente: identifica o endereço mais completo
    baseado no critério de "contém".
-   Parsing estruturado em colunas:
    -   TIPO
    -   LOGRADOURO
    -   NÚMERO
    -   COMPLEMENTO
    -   BAIRRO
    -   CIDADE
    -   ESTADO
    -   PAÍS
-   Geração de um arquivo Excel final formatado.

------------------------------------------------------------------------

## 🏗️ Principais Decisões de Arquitetura

### ✔ Normalização via Regex

Um dicionário de padrões (`REPLACEMENTS`) aplica conversões padronizadas
para abreviações comuns.

### ✔ Extração de coordenadas

A função `extrair_geo()` remove informações no formato `(lat, lon)` e
armazena separadamente.

### ✔ Deduplicação baseada em "endereço contém endereço"

Comparações de substring foram escolhidas por serem mais simples e
eficientes para o tipo de dado.

### ✔ Parsing estruturado

A função `parse_canonical_address()` decompõe o endereço canônico em
partes usando regras baseadas em tokens e padrões predefinidos.

### ✔ Logs detalhados

O script gera logs tanto no console quanto no arquivo
`log_enderecos.txt`.

------------------------------------------------------------------------

## 🧪 Como executar os testes

O script contém testes básicos integrados. Para rodá-los, execute:

``` bash
python3 script.py --test
```

ou

``` bash
py script.py --test
```

Isso executará a função `run_tests()` e validará:

-   Normalização
-   Substituição de vírgulas duplicadas
-   Extração de coordenadas
-   Parsing de endereço simples

------------------------------------------------------------------------

## 🐍 Problemas com venv ou dependências?

Se houver qualquer erro relacionado a módulos ausentes (como `pandas`,
`openpyxl`, etc.), utilize o arquivo:

    requirements.txt

Para instalar todas as dependências:

``` bash
pip install -r requirements.txt
```

Isso garante que o ambiente estará completo e compatível com o script.

------------------------------------------------------------------------
