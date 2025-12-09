package workbook

import (
	"os"
	"path/filepath"
	"strings"
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

func TestSetAndClearCell(t *testing.T) {
	wb := &Workbook{}
	wb.SetCell(3, 2, "hello")
	if got := wb.Cell(3, 2); got != "hello" {
		t.Fatalf("expected hello, got %s", got)
	}
	wb.ClearCell(3, 2)
	if got := wb.Cell(3, 2); got != "" {
		t.Fatalf("expected empty after clear, got %s", got)
	}
}

func TestInsertDeleteRow(t *testing.T) {
	wb := &Workbook{
		Cells:  [][]string{{"a", "b"}, {"c", "d"}},
		Styles: map[string]CellStyle{"A1": {Bold: true}, "B2": {Italic: true}},
		Sheet:  "Sheet1",
	}
	wb.InsertRow(2, []string{"x", "y"})
	if val := wb.Cell(2, 1); val != "x" {
		t.Fatalf("expected inserted row value, got %s", val)
	}
	if _, ok := wb.Styles["B3"]; !ok {
		t.Fatalf("expected style to shift after insert")
	}
	row, ok := wb.DeleteRow(2)
	if !ok || row[0] != "x" {
		t.Fatalf("expected delete row to return inserted data")
	}
	if _, ok := wb.Styles["B2"]; !ok {
		t.Fatalf("expected shifted style to move back after delete")
	}
}

func TestSaveCSV(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "out.csv")
	wb := &Workbook{
		Cells: [][]string{{"a", "b"}, {"c", "d"}},
		Sheet: "Sheet1",
	}
	if err := wb.Save(path); err != nil {
		t.Fatalf("save csv failed: %v", err)
	}
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatalf("read csv: %v", err)
	}
	if got := string(data); !strings.Contains(got, "a,b") {
		t.Fatalf("unexpected csv contents: %s", got)
	}
}

func TestAddSheet(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "book.xlsx")
	wb := SampleWorkbook()
	if err := wb.Save(path); err != nil {
		t.Fatalf("initial save failed: %v", err)
	}
	if err := AddSheet(path, "CopySheet", "Sheet1"); err != nil {
		t.Fatalf("copy sheet failed: %v", err)
	}
	if err := AddSheet(path, "BlankSheet", ""); err != nil {
		t.Fatalf("blank sheet failed: %v", err)
	}
	_, sheets, _, err := FromXLSX(path, "CopySheet")
	if err != nil {
		t.Fatalf("reload failed: %v", err)
	}
	found := 0
	for _, s := range sheets {
		if s == "CopySheet" || s == "BlankSheet" {
			found++
		}
	}
	if found != 2 {
		t.Fatalf("expected new sheets in workbook, got %v", sheets)
	}
}

func TestInsertDeleteColumn(t *testing.T) {
	wb := SampleWorkbook()
	wb.InsertColumn(2)
	if cols := len(wb.Cells[0]); cols != 4 {
		t.Fatalf("expected header row to grow to 4 columns, got %d", cols)
	}
	if val := wb.Cells[1][1]; val != "" {
		t.Fatalf("expected blank cell after column insert, got %q", val)
	}
	col, ok := wb.DeleteColumn(2)
	if !ok {
		t.Fatalf("expected delete column to succeed")
	}
	if len(col) != len(wb.Cells) {
		t.Fatalf("expected removed column slice to match row count")
	}
	if cols := len(wb.Cells[0]); cols != 3 {
		t.Fatalf("expected header row back to 3 columns, got %d", cols)
	}
}
