package main

import (
	"log"
	"net/http"

	"github.com/mhshajib/retroxl"
)

func handler(w http.ResponseWriter, r *http.Request) {
	headers := []string{"A", "B"}
	rows := [][]any{{"x", 1}, {"y", 2}}
	sheet := retroxl.FromRows("Demo", headers, rows)

	w.Header().Set("Content-Type", "application/vnd.ms-excel")
	w.Header().Set("Content-Disposition", `attachment; filename="demo.xls"`)

	if err := retroxl.WriteXLSWriter(w, []retroxl.Sheet{sheet}); err != nil {
		http.Error(w, "failed", 500)
	}
}

func main() {
	http.HandleFunc("/demo.xls", handler)
	log.Fatal(http.ListenAndServe(":8080", nil))
}
