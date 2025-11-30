package retroxl

// Sheet represents a single worksheet in a workbook.
// Headers are optional and, when present, are written as the first row.
// Each row must have the same length as Headers, if Headers are set.
type Sheet struct {
    Name    string
    Headers []string
    Rows    [][]any
}

// FromRows constructs a Sheet from a name, optional headers, and row data.
func FromRows(name string, headers []string, rows [][]any) Sheet {
    return Sheet{
        Name:    name,
        Headers: headers,
        Rows:    rows,
    }
}
