package retroxl

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
)

const (
	nsWorkbook = "urn:schemas-microsoft-com:office:spreadsheet"
	nsOffice   = "urn:schemas-microsoft-com:office:office"
	nsExcel    = "urn:schemas-microsoft-com:office:excel"
	nsSS       = "urn:schemas-microsoft-com:office:spreadsheet"
	nsHTML     = "http://www.w3.org/TR/REC-html40"
)

// WriteXLSFile writes one or more sheets to a file path as
// XML Spreadsheet 2003 content with a .xls extension.
func WriteXLSFile(outPath string, sheets []Sheet) error {
	data, err := WriteXLSBytes(sheets)
	if err != nil {
		return err
	}
	return os.WriteFile(outPath, data, 0o644)
}

// WriteXLSBytes writes one or more sheets and returns the XLS file content
// as a byte slice.
func WriteXLSBytes(sheets []Sheet) ([]byte, error) {
	var buf bytes.Buffer
	if err := WriteXLSWriter(&buf, sheets); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

// WriteXLSWriter writes one or more sheets to the provided io.Writer as
// XML Spreadsheet 2003 content.
func WriteXLSWriter(w io.Writer, sheets []Sheet) error {
	return writeXMLWorkbook(w, sheets)
}

func writeXMLWorkbook(w io.Writer, sheets []Sheet) error {
	if _, err := io.WriteString(w, `<?xml version="1.0"?>`+"\n"); err != nil {
		return err
	}
	if _, err := io.WriteString(w, `<?mso-application progid="Excel.Sheet"?>`+"\n"); err != nil {
		return err
	}

	enc := xml.NewEncoder(w)
	enc.Indent("", "  ")

	wbStart := xml.StartElement{
		Name: xml.Name{Local: "Workbook"},
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "xmlns"}, Value: nsWorkbook},
			{Name: xml.Name{Local: "xmlns:o"}, Value: nsOffice},
			{Name: xml.Name{Local: "xmlns:x"}, Value: nsExcel},
			{Name: xml.Name{Local: "xmlns:ss"}, Value: nsSS},
			{Name: xml.Name{Local: "xmlns:html"}, Value: nsHTML},
		},
	}

	if err := enc.EncodeToken(wbStart); err != nil {
		return err
	}

	if err := writeDefaultStyles(enc); err != nil {
		return err
	}

	for _, s := range sheets {
		if err := writeWorksheet(enc, s); err != nil {
			return err
		}
	}

	if err := enc.EncodeToken(wbStart.End()); err != nil {
		return err
	}

	return enc.Flush()
}

func writeDefaultStyles(enc *xml.Encoder) error {
	stylesStart := xml.StartElement{Name: xml.Name{Local: "Styles"}}
	if err := enc.EncodeToken(stylesStart); err != nil {
		return err
	}

	styleStart := xml.StartElement{
		Name: xml.Name{Local: "Style"},
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "ss:ID"}, Value: "Default"},
		},
	}

	if err := enc.EncodeToken(styleStart); err != nil {
		return err
	}
	if err := enc.EncodeToken(styleStart.End()); err != nil {
		return err
	}

	return enc.EncodeToken(stylesStart.End())
}

func writeWorksheet(enc *xml.Encoder, s Sheet) error {
	wsStart := xml.StartElement{
		Name: xml.Name{Local: "Worksheet"},
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "ss:Name"}, Value: safeSheetName(s.Name)},
		},
	}

	if err := enc.EncodeToken(wsStart); err != nil {
		return err
	}

	tableStart := xml.StartElement{Name: xml.Name{Local: "Table"}}
	if err := enc.EncodeToken(tableStart); err != nil {
		return err
	}

	if len(s.Headers) > 0 {
		if err := writeRow(enc, stringSliceToAny(s.Headers)); err != nil {
			return err
		}
	}

	for _, r := range s.Rows {
		if err := writeRow(enc, r); err != nil {
			return err
		}
	}

	if err := enc.EncodeToken(tableStart.End()); err != nil {
		return err
	}

	return enc.EncodeToken(wsStart.End())
}

func writeRow(enc *xml.Encoder, row []any) error {
	rowStart := xml.StartElement{Name: xml.Name{Local: "Row"}}
	if err := enc.EncodeToken(rowStart); err != nil {
		return err
	}

	for _, v := range row {
		if err := writeCell(enc, v); err != nil {
			return err
		}
	}

	return enc.EncodeToken(rowStart.End())
}

func writeCell(enc *xml.Encoder, v any) error {
	cellStart := xml.StartElement{Name: xml.Name{Local: "Cell"}}
	if err := enc.EncodeToken(cellStart); err != nil {
		return err
	}

	if v == nil {
		if err := enc.EncodeToken(cellStart.End()); err != nil {
			return err
		}
		return nil
	}

	dataStart := xml.StartElement{Name: xml.Name{Local: "Data"}}

	switch val := v.(type) {
	case int, int8, int16, int32, int64,
		uint, uint8, uint16, uint32, uint64,
		float32, float64:
		dataStart.Attr = []xml.Attr{
			{Name: xml.Name{Local: "ss:Type"}, Value: "Number"},
		}
		if err := enc.EncodeToken(dataStart); err != nil {
			return err
		}
		if err := enc.EncodeToken(xml.CharData([]byte(fmt.Sprint(val)))); err != nil {
			return err
		}
		if err := enc.EncodeToken(dataStart.End()); err != nil {
			return err
		}
	case bool:
		dataStart.Attr = []xml.Attr{
			{Name: xml.Name{Local: "ss:Type"}, Value: "Boolean"},
		}
		if err := enc.EncodeToken(dataStart); err != nil {
			return err
		}
		var b string
		if val {
			b = "1"
		} else {
			b = "0"
		}
		if err := enc.EncodeToken(xml.CharData([]byte(b))); err != nil {
			return err
		}
		if err := enc.EncodeToken(dataStart.End()); err != nil {
			return err
		}
	default:
		dataStart.Attr = []xml.Attr{
			{Name: xml.Name{Local: "ss:Type"}, Value: "String"},
		}
		if err := enc.EncodeToken(dataStart); err != nil {
			return err
		}
		if err := enc.EncodeToken(xml.CharData([]byte(fmt.Sprint(val)))); err != nil {
			return err
		}
		if err := enc.EncodeToken(dataStart.End()); err != nil {
			return err
		}
	}

	if err := enc.EncodeToken(cellStart.End()); err != nil {
		return err
	}

	return nil
}

func safeSheetName(s string) string {
	if s == "" {
		return "Sheet1"
	}
	const maxLen = 31
	if len(s) > maxLen {
		s = s[:maxLen]
	}
	return s
}

func stringSliceToAny(in []string) []any {
	out := make([]any, len(in))
	for i, v := range in {
		out[i] = v
	}
	return out
}
