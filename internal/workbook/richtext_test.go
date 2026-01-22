package workbook

import (
	"os"
	"path/filepath"
	"testing"

	excelize "github.com/xuri/excelize/v2"
)

func TestScanRichTextDetectsRuns(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "richtext.xlsx")

	f := excelize.NewFile()
	sheet := f.GetSheetName(f.GetActiveSheetIndex())
	runs := []excelize.RichTextRun{
		{
			Text: "Bold",
			Font: &excelize.Font{Bold: true},
		},
		{
			Text: " Plain",
			Font: &excelize.Font{},
		},
	}
	if err := f.SetCellRichText(sheet, "A1", runs); err != nil {
		t.Fatalf("set rich text: %v", err)
	}
	if err := f.SetCellValue(sheet, "B2", "plain"); err != nil {
		t.Fatalf("set plain cell: %v", err)
	}
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("save xlsx: %v", err)
	}
	if err := f.Close(); err != nil {
		t.Fatalf("close xlsx: %v", err)
	}

	warnings, err := ScanRichText(path, "")
	if err != nil {
		t.Fatalf("scan rich text: %v", err)
	}
	if len(warnings) != 1 {
		t.Fatalf("expected 1 warning, got %d", len(warnings))
	}
	if warnings[0].Sheet != sheet || warnings[0].Cell != "A1" {
		t.Fatalf("unexpected warning %v", warnings[0])
	}
	if warnings[0].Runs < 2 {
		t.Fatalf("expected multiple runs, got %d", warnings[0].Runs)
	}
}

func TestScanRichTextMissingFile(t *testing.T) {
	_, err := ScanRichText(filepath.Join(os.TempDir(), "missing.xlsx"), "")
	if err == nil {
		t.Fatal("expected error for missing file")
	}
}
