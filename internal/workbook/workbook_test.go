package workbook

import (
	"os"
	"path/filepath"
	"testing"
)

func TestSampleWorkbook(t *testing.T) {
	wb := SampleWorkbook()
	if wb.Name != "sample" {
		t.Fatalf("expected name sample got %s", wb.Name)
	}
	if rows, cols := wb.MaxCoords(); rows != 5 || cols != 3 {
		t.Fatalf("expected 5 rows and 3 cols, got %d, %d", rows, cols)
	}
	if val := wb.Cell(2, 1); val != "Foam" {
		t.Fatalf("expected Foam in row2 col1, got %s", val)
	}
}

func TestCellOutOfBounds(t *testing.T) {
	wb := &Workbook{Cells: [][]string{{"a"}}}
	for _, tc := range [][2]int{{0, 0}, {1, 2}, {2, 1}} {
		if val := wb.Cell(tc[0], tc[1]); val != "" {
			t.Fatalf("expected empty for row %d col %d", tc[0], tc[1])
		}
	}
}

func TestFromCSV(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "sample.csv")
	content := "name,qty\nwidget,3\n"
	if err := os.WriteFile(path, []byte(content), 0o644); err != nil {
		t.Fatalf("write csv: %v", err)
	}

	wb, err := FromCSV(path)
	if err != nil {
		t.Fatalf("FromCSV failed: %v", err)
	}
	if got := wb.Cell(2, 2); got != "3" {
		t.Fatalf("expected qty 3 got %s", got)
	}
}

func TestColumnName(t *testing.T) {
	cases := map[int]string{
		1:   "A",
		26:  "Z",
		27:  "AA",
		52:  "AZ",
		53:  "BA",
		703: "AAA",
		0:   "A",
		-1:  "A",
	}
	for col, want := range cases {
		if got := ColumnName(col); got != want {
			t.Fatalf("col %d: expected %s, got %s", col, want, got)
		}
	}
}
