# MultiXLS2CSV

A small tool made because Microsoft Excel doesn't support exporting to csv all sheets at once from
an Excel file.

## Usage/Examples

```
Usage:
       MultiXLS2CSV.exe [flags] <xlsx-to-be-read>
  -delimiter string
        Delimiter to use between fields in the CSV (default ",")
  -output string
        Path to the output folder where the CSVs will be saved (default ".")
```

```sh
MultiXLS2CSV.exe accounting.xlsx
```

```sh
MultiXLS2CSV.exe -delimiter ";" -output "csv/accounting" accounting.xlsx
```

## Run from source

```sh
go run main.go -delimiter ";" -output "csv/accounting" accounting.xlsx
```
