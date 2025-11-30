// Package retroxl is a pure-Go library for generating legacy-compatible XLS
// files from modern spreadsheet inputs. It converts XLSX, CSV, TSV, and
// in-memory tabular data into SpreadsheetML-based XLS files that open in
// Microsoft Excel and pass validation in legacy systems.
//
// RetroXL is designed for integrations where banks, government portals,
// or older enterprise systems still require `.xls` uploads instead of modern
// `.xlsx`. The library has no external dependencies and does not rely on
// LibreOffice, Python, or system binaries.
//
// The API supports converting files from disk or constructing sheets in
// memory. Output can be written directly to disk, returned as []byte,
// or streamed to any io.Writer (HTTP response, gRPC, S3 uploads, etc.).
package retroxl
