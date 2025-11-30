package retroxl

import (
    "os"
    "testing"
)

func TestWriteXLSFile(t *testing.T) {
    headers := []string{"A", "B"}
    rows := [][]any{
        {1, "x"},
        {2, "y"},
    }

    sheet := FromRows("Test", headers, rows)
    out := "test_output.xls"
    defer func() {
        _ = os.Remove(out)
    }()

    if err := WriteXLSFile(out, []Sheet{sheet}); err != nil {
        t.Fatalf("WriteXLSFile error: %v", err)
    }

    if _, err := os.Stat(out); err != nil {
        t.Fatalf("expected output file to exist: %v", err)
    }
}
