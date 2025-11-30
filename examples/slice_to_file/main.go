package main

import (
	"log"

	"github.com/mhshajib/retroxl"
)

type Payment struct {
	AccountNo string
	Amount    float64
	Reference string
}

func main() {
	items := []Payment{
		{"1234567890", 1200.50, "Invoice-001"},
		{"0987654321", 300.00, "Invoice-002"},
	}

	headers := []string{"AccountNo", "Amount", "Reference"}
	var rows [][]any
	for _, p := range items {
		rows = append(rows, []any{p.AccountNo, p.Amount, p.Reference})
	}

	sheet := retroxl.FromRows("Payments", headers, rows)
	if err := retroxl.WriteXLSFile("payments.xls", []retroxl.Sheet{sheet}); err != nil {
		log.Fatal(err)
	}
}
