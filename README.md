# TaxesHelper
A tiny command-line program that makes it easier to pay taxes in Bulgaria.

This program is intended to be used by people that trade stocks in etrade.

# Features
## Holdings Analyzer
The program can help you fill Application 8.

```bash
$ dotnet TaxesHelper.dll holdings -h
holdings:
  This command generates the output for Application №8.

Usage:
  TaxesHelper holdings [options]

Options:
  --file <file> (REQUIRED)    Path to the Etrade's excel file for sellable stocks.
                              Go to Etrade > Stock Plan > View By Status > Download Expanded and save the file locally.
  --interactive               If true the program prints 1 line and waits for user input before printing the next.
                              Useful if you're filling the application. If false (default), dumps the whole output in one go.
  -?, -h, --help              Show help and usage information
```

## Sells Analyzer
```bash
$ dotnet TaxesHelper.dll sells -h
sells:
  This command generates the output for Application №5.

Usage:
  TaxesHelper sells [options]

Options:
  --file <file> (REQUIRED)                  Path to the Etrade's excel file for sellable stocks.
                                            Go to Etrade > My Account > Gains & Losses > Download Expanded and save the file locally.
                                            Note: you must filter by year!
  --sell-prices <sell-prices> (REQUIRED)    List of the sell order prices that are in the provided file.
                                            Etrade provides an 'Adjusted price' which includes their commission.
                                            To find out your sell prices, go to Etrade > My Account > Orders > Execution price.
                                            Note: each price must be provided as a separate argument e.g. --sell-prices 100 120
  -?, -h, --help                            Show help and usage information

```
