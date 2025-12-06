package workbook

import (
	"encoding/csv"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
)

// Workbook is a minimal in-memory representation for demo purposes.
type Workbook struct {
	Cells [][]string
	Name  string
	Sheet string
}

// SampleWorkbook seeds demo data without hitting the filesystem.
func SampleWorkbook() *Workbook {
	cells := [][]string{
		{"Item", "Qty", "Price"},
		{"Foam", "2", "$15"},
		{"Brush", "5", "$7"},
		{"Ink", "1", "$42"},
		{"Total", "8", "$64"},
	}
	return &Workbook{Cells: cells, Name: "sample", Sheet: "Sheet1"}
}

// FromFile loads either CSV or XLSX data into a Workbook.
func FromFile(path, sheet string) (*Workbook, error) {
	if path == "" {
		return nil, fmt.Errorf("path is required")
	}
	switch strings.ToLower(filepath.Ext(path)) {
	case ".csv":
		return FromCSV(path)
	case ".xlsx":
		return FromXLSX(path, sheet)
	default:
		return nil, fmt.Errorf("unsupported extension %s", filepath.Ext(path))
	}
}

// FromCSV loads a CSV file into a Workbook; it uses comma delimiter for now.
func FromCSV(path string) (*Workbook, error) {
	file, err := os.Open(path)
	if err != nil {
		return nil, fmt.Errorf("open csv: %w", err)
	}
	defer file.Close()

	reader := csv.NewReader(file)
	var rows [][]string
	for {
		record, err := reader.Read()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("read csv: %w", err)
		}
		rows = append(rows, record)
	}
	return &Workbook{Cells: rows, Name: path, Sheet: "Sheet1"}, nil
}

// Cell returns the value at 1-based row/col, empty string if out of bounds.
func (w *Workbook) Cell(row, col int) string {
	if w == nil || row < 1 || col < 1 {
		return ""
	}
	rowIdx := row - 1
	colIdx := col - 1
	if rowIdx >= len(w.Cells) {
		return ""
	}
	rowData := w.Cells[rowIdx]
	if colIdx >= len(rowData) {
		return ""
	}
	return rowData[colIdx]
}

// MaxCoords returns the max row and column counts.
func (w *Workbook) MaxCoords() (int, int) {
	rows := len(w.Cells)
	cols := 0
	for _, r := range w.Cells {
		if len(r) > cols {
			cols = len(r)
		}
	}
	return rows, cols
}
