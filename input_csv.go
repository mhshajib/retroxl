package retroxl

import (
	"encoding/csv"
	"os"
)

// CSVToSheets reads a delimited text file (CSV or TSV) from the provided path
// and converts it into a single Sheet. The first row is treated as headers.
func CSVToSheets(path string, sep rune) ([]Sheet, error) {
	f, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	defer func() {
		_ = f.Close()
	}()

	r := csv.NewReader(f)
	r.Comma = sep

	raw, err := r.ReadAll()
	if err != nil {
		return nil, err
	}

	if len(raw) == 0 {
		return []Sheet{
			FromRows("Sheet1", nil, nil),
		}, nil
	}

	headers := make([]string, len(raw[0]))
	for i, v := range raw[0] {
		headers[i] = v
	}

	var rows [][]any
	for _, rr := range raw[1:] {
		row := make([]any, len(rr))
		for i, v := range rr {
			row[i] = v
		}
		rows = append(rows, row)
	}

	sheet := FromRows("Sheet1", headers, rows)
	return []Sheet{sheet}, nil
}
