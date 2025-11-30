// Package retroxl provides utilities for generating legacy-compatible XLS files
// from modern spreadsheet inputs such as XLSX, CSV, TSV, and in-memory data.
//
// The package encodes data as XML Spreadsheet 2003 (SpreadsheetML) and writes
// it with a .xls extension. The generated files open in Excel and are suitable
// for legacy upload systems such as banking or government portals that still
// require .xls uploads.
package retroxl
