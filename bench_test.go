package retroxl

import "testing"

func BenchmarkWriteXLSBytes(b *testing.B) {
    headers := []string{"Col1", "Col2", "Col3"}
    var rows [][]any
    for i := 0; i < 1000; i++ {
        rows = append(rows, []any{i, float64(i) * 1.1, "row"})
    }

    sheet := FromRows("Sheet1", headers, rows)
    sheets := []Sheet{sheet}

    b.ResetTimer()

    for i := 0; i < b.N; i++ {
        if _, err := WriteXLSBytes(sheets); err != nil {
            b.Fatalf("WriteXLSBytes error: %v", err)
        }
    }
}
