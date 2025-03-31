<p align="right">
  <a href="https://github.com/guialmeidan/dividends_extractor/blob/main/README-pt.md">
    <img src="https://img.shields.io/badge/PORTUGUESE-4285F4?style=flat&logo=googletranslate&logoColor=white" alt="Google Translate Badge">
  </a>
</p>
# Dividend Extractor

## Description

This code reads a Google Spreadsheets file containing real estate investment funds listed on B3 (São Paulo Stock Exchange) and returns an Excel file with the dividend values for a specified period.
The code is recommended for shareholders who hold shares in Real Estate Investment Funds and wish to automate the control of their dividends. 

## Requirements

- You must have a spreadsheet hosted on Google Spreadsheets as per the "REITs" file in the "template" folder.
- You must have a "credentials.json" file generated for authentication. See instructions below.
- Check the library dependencies in the "pyproject.toml" file.

## Template
This is the structure of the "Template" file, which must be hosted on Google Spreadsheets:
![](https://github.com/guialmeidan/dividends_extractor/blob/main/images/template_google_spreadsheets.png?raw=true)

Columns A (Ticker) and E (Shares) are the only ones that must be preserved for the code to work; however, you can change their positions as long as corresponding changes are made in the code.

- **Ticker**: A _string_ field containing the code of the real estate investment fund in the format "XXXX11".
If the template changes, you must modify where `row[0]` is used in lines 174 and 178 to the number of the column corresponding to the new layout (see below).

- **Shares**: An _int_ field containing the total shares the shareholder holds in the real estate investment fund.
If the template changes, you must modify where `row[4]` is used in line 175 to the number of the column corresponding to the new layout (see below).
    ```sh
    for row in rows[1:]:
        # Checks if a fund is registered in the sheet
        if row[0]:
            if int(row[4]) > 0: # Proceeds with extraction only if shares are available
                # Adds the '.SA' prefix to refer to the São Paulo Stock Exchange - Brazil
                ticker = row[0] + ".SA"
    ```

The name of the spreadsheet (REITs) and the tab where the information is located (Portfolio) are also important and can be modified on lines 163 and 166:

```sh
# Searching for the spreadsheet by name
spreadsheet = client.open("REITs")

# Selecting the 'Portfolio' tab for reading
sheet = spreadsheet.worksheet("Portfolio")
  ```

### Credentials File

The `"credentials.json"` file must be located inside the `"src\dividend_extractor\credentials"` folder. For security reasons, this folder with the file is not available in this repository. The user must create the folder and upload the file.
Instructions on how to create the file with credentials are available at this link: [Create access credentials | Google Workspace](https://developers.google.com/workspace/guides/create-credentials)

## Execution
To run the code, simply modify the Start Date and End Date, specified in lines 143 and 147. The accepted format is `dd/mm/aaaa`:
```sh
# Defines the start date for dividend search
start_date = "01/03/2025"
start_date = extract_date(start_date)

# Defines the end date for dividend search
end_date = "31/03/2025"
```
## Output File

The output file `Dividends.xlsx` is formatted as follows:

![](https://github.com/guialmeidan/dividends_extractor/blob/main/images/output_image.png?raw=true)

- **Date**: The dividend payment date
- **Ticker**: The ticker name
- **Dividend**: Total dividends received in the period, already multiplied by the total shares
- **Shares**: Total shares the shareholder holds
