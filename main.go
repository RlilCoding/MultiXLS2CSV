package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"os"
	"path/filepath"

	"github.com/xuri/excelize/v2"
)

var (
	outputFolder = flag.String("output", ".", "Path to the output folder where the CSVs will be saved")
	csvDelimiter = flag.String("delimiter", ",", "Delimiter to use between fields in the CSV")
)

func writeSheetToCSV(xlsx *excelize.File, sheetName string, delimiter rune, outputFolder string) error {
	// get all the rows in the sheet
	rows, err := xlsx.GetRows(sheetName)
	if err != nil {
		return fmt.Errorf("error reading sheet %s: %v", sheetName, err)
	}

	// create the output file
	outputFile, err := os.Create(filepath.Join(outputFolder, sheetName+".csv"))
	if err != nil {
		return fmt.Errorf("error creating output file: %v", err)
	}
	defer outputFile.Close()

	// make a CSV writer
	csvWriter := csv.NewWriter(outputFile)
	csvWriter.Comma = delimiter

	// get the maximum number of columns in a row
	maxColumns := 0
	for _, row := range rows {
		maxColumns = max(maxColumns, len(row))
	}
	
	for _, row := range rows {
		// pad the row with empty strings if it has less columns than the maximum
		for len(row) < maxColumns {
			row = append(row, "")
		}
		if err := csvWriter.Write(row); err != nil {
			return fmt.Errorf("error writing row to CSV: %v", err)
		}
	}

	// flush the CSV writer
	csvWriter.Flush()
	if err := csvWriter.Error(); err != nil {
		return fmt.Errorf("error while writing CSV to file: %v", err)
	}

	return nil
}


func main() {
	flag.Usage = func() {
		fmt.Fprintf(os.Stderr, `
Reads an Excel file and writes every sheet to csv files in the output folder.

Usage:
	%s [flags] <excel-to-be-read>
`, os.Args[0])
		flag.PrintDefaults()
	}

	flag.Parse()

	var delimiter rune
	if len(*csvDelimiter) != 1 {
		fmt.Fprintf(os.Stderr, "Delimiter must be a single character\n")
		os.Exit(1)
	}
	delimiter = rune((*csvDelimiter)[0])

	if err := os.MkdirAll(*outputFolder, 0755); err != nil {
		fmt.Fprintf(os.Stderr, "Error creating output folder: %v\n", err)
		os.Exit(1)
	}

	if flag.NArg() != 1 {
		flag.Usage()
		os.Exit(1)
	}

	xlsx, err := excelize.OpenFile(flag.Arg(0))
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error opening file: %v\n", err)
		os.Exit(1)
	}

	// iterate over the sheets
	for _, sheetName := range xlsx.GetSheetMap() {
		if err := writeSheetToCSV(xlsx, sheetName, delimiter, *outputFolder); err != nil {
			fmt.Fprintf(os.Stderr, "Error writing sheet to CSV: %v\n", err)
			os.Exit(1)
		}
	}

}
