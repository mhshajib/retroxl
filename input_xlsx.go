package retroxl

import (
    "fmt"

    "github.com/xuri/excelize/v2"
)

// XLSXToSheets reads an .xlsx file from the provided path and converts it
// into a slice of Sheet.
func XLSXToSheets(path string) ([]Sheet, error) {
    f, err := excelize.OpenFile(path)
    if err != nil {
        return nil, fmt.Errorf("open xlsx: %w", err)
    }
    defer func() {
        _ = f.Close()
    }()

    var result []Sheet

    sheetNames := f.GetSheetList()
    for _, sn := range sheetNames {
        rows, err := f.Rows(sn)
        if err != nil {
            return nil, fmt.Errorf("read rows: %w", err)
        }

        var allRows [][]any
        for rows.Next() {
            cols, err := rows.Columns()
            if err != nil {
                return nil, fmt.Errorf("read row columns: %w", err)
            }
            row := make([]any, len(cols))
            for i, c := range cols {
                row[i] = c
            }
            allRows = append(allRows, row)
        }

        var headers []string
        if len(allRows) > 0 {
            headers = make([]string, len(allRows[0]))
            for i, v := range allRows[0] {
                if s, ok := v.(string); ok {
                    headers[i] = s
                } else {
                    headers[i] = fmt.Sprint(v)
                }
            }
            allRows = allRows[1:]
        }

        result = append(result, Sheet{
            Name:    sn,
            Headers: headers,
            Rows:    allRows,
        })
    }

    return result, nil
}
