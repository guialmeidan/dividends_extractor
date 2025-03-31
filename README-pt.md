<p align="right">
  <a href="https://github.com/guialmeidan/dividends_extractor/blob/main/README.md">
    <img src="https://img.shields.io/badge/ENGLISH-4285F4?style=flat&logo=googletranslate&logoColor=white" alt="Google Translate Badge">
  </a>
</p>
# Extrator de Dividendos

## Descrição

Este código faz a leitura de um arquivo do Google Spreadsheets que contém fundos imobiliários listados na B3 (Bolsa de Valores de São Paulo) e retorna um arquivo do Excel com os valores dos dividendos para um período especificado.
O código é recomendado para acionistas que possuem cotas de Fundos Imobiliários e desejam automatizar o controle de seus dividendos. 

## Requisitos

- Possuir uma planilha hospedada no Google Spreadsheets conforme o arquivo "REITs" da pasta "template".
- Ter um arquivo "credentials.json" gerado, para autenticação da leitura. Ver orientações abaixo.
- Consultar dependências de bibliotecas no arquivo "pyproject.toml".

## Template
Esta é a estrutura do arquivo _"Template"_, que deve ser hospedada no Google Spreadsheets:
![](https://github.com/guialmeidan/dividends_extractor/blob/main/images/template_google_spreadsheets.png?raw=true)

As colunas A (Ticker) e E (Shares) são únicas que devem ser preservadas para funcionamento do código, entretanto é possível alterá-las de posição, desde que alterações sejam realizadas também no código.

- **Ticker**: Campo do tipo _string_, que contém o código do fundo imobiliário no formato "XXXX11".
    Caso haja alteração do template, é necessário alterar onde se lê `row[0]` nas linhas 174 e 178 para o número da coluna correspondente no novo layout (ver abaixo).

- **Shares**: Campo do tipo _int_, que contém o total de cotas do fundo imobiliário que o acionista possui.
Caso haja alteração do template, é necessário alterar onde se lê `row[4]` nas linhas 175 para o número da coluna correspondente no novo layout (ver abaixo).
    ```sh
    for row in rows[1:]:
        # Checks if a fund is registered in the sheet
        if row[0]:
            if int(row[4]) > 0: # Proceeds with extraction only if shares are available
                # Adds the '.SA' prefix to refer to the São Paulo Stock Exchange - Brazil
                ticker = row[0] + ".SA"
    ```

O nome da planilha (REITs) e da aba onde estão as informações (Portfolio) também são importantes e podem ser modificadas nas linhas 163 e 166:

```sh
# Searching for the spreadsheet by name
spreadsheet = client.open("REITs")

# Selecting the 'Portfolio' tab for reading
sheet = spreadsheet.worksheet("Portfolio")
  ```

### Arquivo Credentials

O arquivo `"credentials.json"` deve estar dentro da pasta `"src\dividend_extractor\credentials"`. Por motivo de segurança esta pasta com o arquivo não está disponível neste repositório. O usuário deverá criar a pasta e fazer upload do arquivo.
Instruções de como criar o arquivo com as credenciais estão disponíveis neste link: [Criar credenciais de acesso | Google Workspace](https://developers.google.com/workspace/guides/create-credentials)

## Execução
Para executar o código, basta alterar a Data Inicial e a Data Final, descritas nas linhas 143 e 147. O formato aceito é `dd/mm/aaaa`:
```sh
# Defines the start date for dividend search
start_date = "01/03/2025"
start_date = extract_date(start_date)

# Defines the end date for dividend search
end_date = "31/03/2025"
```
## Arquivo de Saída

O arquivo de saída `Dividends.xlsx` é formatado desta forma:

![](https://github.com/guialmeidan/dividends_extractor/blob/main/images/output_image.png?raw=true)

- **Date**: Data de pagamento do dividendo
- **Ticker**: Nome do ticker
- **Dividend**: Total de dividendos recebidos no período, já multiplicado pelo total de cotas
- **Shares**: Total de cotas que o acionista possui
