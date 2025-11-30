package retroxl

import (
	"fmt"
	"path/filepath"
	"strings"
)

// PathToSheets reads a file from disk and converts it into a slice of Sheet.
// The conversion is based on the file extension.
//
// Supported extensions:
//
//	.xlsx
//	.csv
//	.tsv
func PathToSheets(inPath string) ([]Sheet, error) {
	ext := strings.ToLower(filepath.Ext(inPath))

	switch ext {
	case ".xlsx":
		return XLSXToSheets(inPath)
	case ".csv":
		return CSVToSheets(inPath, ',')
	case ".tsv":
		return CSVToSheets(inPath, '\t')
	default:
		return nil, fmt.Errorf("unsupported input format: %s", ext)
	}
}

// ConvertXLSXToXLSFile converts an .xlsx file on disk to an .xls file on disk.
func ConvertXLSXToXLSFile(inPath, outPath string) error {
	sheets, err := XLSXToSheets(inPath)
	if err != nil {
		return err
	}
	return WriteXLSFile(outPath, sheets)
}

// ConvertXLSXToXLSBytes converts an .xlsx file on disk to XLS content returned
// as a byte slice.
func ConvertXLSXToXLSBytes(inPath string) ([]byte, error) {
	sheets, err := XLSXToSheets(inPath)
	if err != nil {
		return nil, err
	}
	return WriteXLSBytes(sheets)
}

// ConvertAnyToXLSFile converts any supported source file on disk to an .xls
// file on disk.
func ConvertAnyToXLSFile(inPath, outPath string) error {
	sheets, err := PathToSheets(inPath)
	if err != nil {
		return err
	}
	return WriteXLSFile(outPath, sheets)
}

// ConvertAnyToXLSBytes converts any supported source file on disk to XLS
// content returned as a byte slice.
func ConvertAnyToXLSBytes(inPath string) ([]byte, error) {
	sheets, err := PathToSheets(inPath)
	if err != nil {
		return nil, err
	}
	return WriteXLSBytes(sheets)
}
