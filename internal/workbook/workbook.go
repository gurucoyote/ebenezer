package workbook

import (
	"encoding/csv"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strconv"
	"strings"
)

// Workbook is a minimal in-memory representation for demo purposes.
type Workbook struct {
	Cells  [][]string
	Name   string
	Sheet  string
	Styles map[string]CellStyle
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
	return &Workbook{Cells: cells, Name: "sample", Sheet: "Sheet1", Styles: map[string]CellStyle{}}
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
	return &Workbook{Cells: rows, Name: path, Sheet: "Sheet1", Styles: map[string]CellStyle{}}, nil
}

// Style returns style information for the given cell address (e.g., "B2").
func (w *Workbook) Style(address string) (CellStyle, bool) {
	if w == nil || w.Styles == nil {
		return CellStyle{}, false
	}
	style, ok := w.Styles[strings.ToUpper(address)]
	return style, ok
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

// Row returns a copy of the row at the given index (1-based).
func (w *Workbook) Row(row int) []string {
	if row < 1 || row > len(w.Cells) {
		return nil
	}
	dup := make([]string, len(w.Cells[row-1]))
	copy(dup, w.Cells[row-1])
	return dup
}

// SetCell writes the value at the provided 1-based row/column, expanding the
// in-memory grid as needed.
func (w *Workbook) SetCell(row, col int, value string) {
	if w == nil || row < 1 || col < 1 {
		return
	}
	w.ensureCell(row, col)
	w.Cells[row-1][col-1] = value
}

// ClearCell blanks the cell if it exists.
func (w *Workbook) ClearCell(row, col int) {
	if w == nil || row < 1 || col < 1 {
		return
	}
	if row-1 >= len(w.Cells) {
		return
	}
	rowData := w.Cells[row-1]
	if col-1 >= len(rowData) {
		return
	}
	rowData[col-1] = ""
}

func (w *Workbook) ensureCell(row, col int) {
	for len(w.Cells) < row {
		w.Cells = append(w.Cells, []string{})
	}
	rowData := w.Cells[row-1]
	if len(rowData) < col {
		rowData = append(rowData, make([]string, col-len(rowData))...)
		w.Cells[row-1] = rowData
	}
}

// ColumnName converts a 1-based column index into its Excel column string.
func ColumnName(col int) string {
	if col <= 0 {
		return "A"
	}
	name := ""
	for col > 0 {
		col--
		name = string(rune('A'+(col%26))) + name
		col /= 26
	}
	return name
}

// InsertRow inserts the provided data before the given row index (1-based).
// If data is nil, a zeroed row is inserted.
func (w *Workbook) InsertRow(idx int, data []string) {
	if idx < 1 {
		idx = 1
	}
	if idx > len(w.Cells)+1 {
		idx = len(w.Cells) + 1
	}
	w.normalizeRow(&data)
	w.Cells = append(w.Cells, nil)
	copy(w.Cells[idx:], w.Cells[idx-1:])
	w.Cells[idx-1] = data
	w.shiftStylesRowsInsert(idx)
}

// DeleteRow removes the row at the given index and returns a copy plus a
// boolean indicating success.
func (w *Workbook) DeleteRow(idx int) ([]string, bool) {
	if idx < 1 || idx > len(w.Cells) {
		return nil, false
	}
	removed := w.Row(idx)
	w.Cells = append(w.Cells[:idx-1], w.Cells[idx:]...)
	w.shiftStylesRowsDelete(idx)
	return removed, true
}

// SetRow overwrites the row at the given index with the provided data,
// expanding as needed.
func (w *Workbook) SetRow(idx int, data []string) {
	if idx < 1 {
		return
	}
	w.normalizeRow(&data)
	for len(w.Cells) < idx {
		w.Cells = append(w.Cells, make([]string, len(data)))
	}
	row := w.Cells[idx-1]
	if len(row) != len(data) {
		row = make([]string, len(data))
	}
	copy(row, data)
	w.Cells[idx-1] = row
}

// SetStyle assigns a style to the provided row/column.
func (w *Workbook) SetStyle(row, col int, style CellStyle) {
	if row < 1 || col < 1 {
		return
	}
	addr := fmt.Sprintf("%s%d", ColumnName(col), row)
	if style.Empty() {
		if w.Styles != nil {
			delete(w.Styles, addr)
		}
		return
	}
	if w.Styles == nil {
		w.Styles = map[string]CellStyle{}
	}
	w.Styles[addr] = style
}

func (w *Workbook) normalizeRow(data *[]string) {
	cols := w.maxCols()
	if cols == 0 {
		cols = len(*data)
	}
	if len(*data) == 0 {
		if cols == 0 {
			*data = []string{}
			return
		}
		*data = make([]string, cols)
		return
	}
	if cols == 0 {
		cols = len(*data)
	}
	if len(*data) < cols {
		*data = append(*data, make([]string, cols-len(*data))...)
	} else if len(*data) > cols && cols > 0 {
		*data = (*data)[:cols]
	}
}

func (w *Workbook) maxCols() int {
	cols := 0
	for _, row := range w.Cells {
		if len(row) > cols {
			cols = len(row)
		}
	}
	return cols
}

func (w *Workbook) shiftStylesRowsInsert(idx int) {
	if w.Styles == nil {
		return
	}
	updated := make(map[string]CellStyle, len(w.Styles))
	for addr, style := range w.Styles {
		col, row, err := splitAddress(addr)
		if err != nil {
			continue
		}
		if row >= idx {
			row++
		}
		updated[fmt.Sprintf("%s%d", col, row)] = style
	}
	w.Styles = updated
}

func (w *Workbook) shiftStylesRowsDelete(idx int) {
	if w.Styles == nil {
		return
	}
	updated := make(map[string]CellStyle, len(w.Styles))
	for addr, style := range w.Styles {
		col, row, err := splitAddress(addr)
		if err != nil {
			continue
		}
		if row == idx {
			continue
		}
		if row > idx {
			row--
		}
		updated[fmt.Sprintf("%s%d", col, row)] = style
	}
	w.Styles = updated
}

func splitAddress(address string) (col string, row int, err error) {
	addr := strings.TrimSpace(strings.ToUpper(address))
	if addr == "" {
		return "", 0, fmt.Errorf("empty address")
	}
	var letters, digits strings.Builder
	for _, r := range addr {
		switch {
		case r >= 'A' && r <= 'Z':
			letters.WriteRune(r)
		case r >= '0' && r <= '9':
			digits.WriteRune(r)
		default:
			return "", 0, fmt.Errorf("invalid address %s", address)
		}
	}
	if letters.Len() == 0 || digits.Len() == 0 {
		return "", 0, fmt.Errorf("invalid address %s", address)
	}
	row, err = strconv.Atoi(digits.String())
	if err != nil {
		return "", 0, err
	}
	return letters.String(), row, nil
}
