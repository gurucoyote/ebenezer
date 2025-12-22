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

// DefaultCSVDelimiter is the rune used when no custom delimiter is provided.
const DefaultCSVDelimiter rune = ';'

// CSVOption configures CSV parsing or serialization behavior.
type CSVOption func(*csvOptions)

type csvOptions struct {
	delimiter rune
}

// WithCSVDelimiter overrides the default rune used to split and join CSV values.
func WithCSVDelimiter(delimiter rune) CSVOption {
	return func(opts *csvOptions) {
		if delimiter != 0 {
			opts.delimiter = delimiter
		}
	}
}

func newCSVOptions(opts []CSVOption) csvOptions {
	cfg := csvOptions{delimiter: DefaultCSVDelimiter}
	for _, opt := range opts {
		if opt == nil {
			continue
		}
		opt(&cfg)
	}
	return cfg
}

// Workbook is a minimal in-memory representation for demo purposes.
type Workbook struct {
	Cells        [][]string
	Name         string
	Sheet        string
	Styles       map[string]CellStyle
	ActiveCell   string
	ColumnWidths map[int]float64
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
	return &Workbook{Cells: cells, Name: "sample", Sheet: "Sheet1", Styles: map[string]CellStyle{}, ActiveCell: "A1", ColumnWidths: map[int]float64{}}
}

// FromFile loads either CSV or XLSX data into a Workbook and returns the sheet
// names plus the workbook's last active cell (when available).
func FromFile(path, sheet string, opts ...CSVOption) (*Workbook, []string, string, error) {
	if path == "" {
		return nil, nil, "", fmt.Errorf("path is required")
	}
	switch strings.ToLower(filepath.Ext(path)) {
	case ".csv":
		wb, err := FromCSV(path, opts...)
		return wb, []string{"Sheet1"}, "", err
	case ".xlsx":
		return FromXLSX(path, sheet)
	default:
		return nil, nil, "", fmt.Errorf("unsupported extension %s", filepath.Ext(path))
	}
}

// Save writes the workbook back to the given path, choosing CSV or XLSX based
// on extension.
func (w *Workbook) Save(path string, opts ...CSVOption) error {
	if w == nil {
		return fmt.Errorf("no workbook data to save")
	}
	csvOpts := newCSVOptions(opts)
	ext := strings.ToLower(filepath.Ext(path))
	switch ext {
	case ".csv":
		return w.saveCSV(path, csvOpts.delimiter)
	case ".xlsx":
		return w.saveXLSX(path)
	default:
		return fmt.Errorf("unsupported extension %s", ext)
	}
}

// FromCSV loads a CSV file into a Workbook using the provided CSV options.
func FromCSV(path string, opts ...CSVOption) (*Workbook, error) {
	file, err := os.Open(path)
	if err != nil {
		return nil, fmt.Errorf("open csv: %w", err)
	}
	defer file.Close()

	csvOpts := newCSVOptions(opts)
	reader := csv.NewReader(file)
	reader.Comma = csvOpts.delimiter
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
	return &Workbook{Cells: rows, Name: path, Sheet: "Sheet1", Styles: map[string]CellStyle{}, ActiveCell: "A1", ColumnWidths: map[int]float64{}}, nil
}

func (w *Workbook) saveCSV(path string, delimiter rune) error {
	file, err := os.Create(path)
	if err != nil {
		return fmt.Errorf("create csv: %w", err)
	}
	defer file.Close()

	writer := csv.NewWriter(file)
	writer.Comma = delimiter
	defer writer.Flush()

	for _, row := range w.Cells {
		if err := writer.Write(row); err != nil {
			return fmt.Errorf("write csv: %w", err)
		}
	}
	return writer.Error()
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

// InsertColumn inserts a blank column before the provided index (1-based).
func (w *Workbook) InsertColumn(idx int) {
	if idx < 1 {
		idx = 1
	}
	maxCols := w.maxCols()
	if maxCols == 0 {
		maxCols = 1
	}
	if idx > maxCols+1 {
		idx = maxCols + 1
	}
	if len(w.Cells) == 0 {
		w.Cells = [][]string{{}}
	}
	for i := range w.Cells {
		row := w.Cells[i]
		if len(row) < idx-1 {
			row = append(row, make([]string, idx-1-len(row))...)
		}
		row = append(row, "")
		copy(row[idx:], row[idx-1:])
		row[idx-1] = ""
		w.Cells[i] = row
	}
	w.shiftStylesColumnsInsert(idx)
	w.shiftColumnWidthsInsert(idx)
}

// DeleteColumn removes the column at the given index and returns a copy
// plus a boolean indicating success.
func (w *Workbook) DeleteColumn(idx int) ([]string, bool) {
	if idx < 1 {
		return nil, false
	}
	maxCols := w.maxCols()
	if idx > maxCols {
		return nil, false
	}
	removed := make([]string, len(w.Cells))
	for i := range w.Cells {
		row := w.Cells[i]
		if idx-1 < len(row) {
			removed[i] = row[idx-1]
			row = append(row[:idx-1], row[idx:]...)
			w.Cells[i] = row
		}
	}
	w.shiftStylesColumnsDelete(idx)
	w.shiftColumnWidthsDelete(idx)
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

func lettersToNumber(s string) int {
	result := 0
	for _, r := range s {
		if r < 'A' || r > 'Z' {
			continue
		}
		result = result*26 + int(r-'A'+1)
	}
	if result == 0 {
		return 1
	}
	return result
}

func (w *Workbook) shiftStylesColumnsInsert(idx int) {
	if w.Styles == nil {
		return
	}
	updated := make(map[string]CellStyle, len(w.Styles))
	for addr, style := range w.Styles {
		colLetters, row, err := splitAddress(addr)
		if err != nil {
			continue
		}
		col := lettersToNumber(colLetters)
		if col >= idx {
			col++
		}
		updated[fmt.Sprintf("%s%d", ColumnName(col), row)] = style
	}
	w.Styles = updated
}

func (w *Workbook) shiftStylesColumnsDelete(idx int) {
	if w.Styles == nil {
		return
	}
	updated := make(map[string]CellStyle, len(w.Styles))
	for addr, style := range w.Styles {
		colLetters, row, err := splitAddress(addr)
		if err != nil {
			continue
		}
		col := lettersToNumber(colLetters)
		if col == idx {
			continue
		}
		if col > idx {
			col--
		}
		updated[fmt.Sprintf("%s%d", ColumnName(col), row)] = style
	}
	w.Styles = updated
}

func (w *Workbook) shiftColumnWidthsInsert(idx int) {
	if w.ColumnWidths == nil {
		return
	}
	updated := make(map[int]float64, len(w.ColumnWidths)+1)
	for col, width := range w.ColumnWidths {
		if col >= idx {
			updated[col+1] = width
			continue
		}
		updated[col] = width
	}
	w.ColumnWidths = updated
}

func (w *Workbook) shiftColumnWidthsDelete(idx int) {
	if w.ColumnWidths == nil {
		return
	}
	updated := make(map[int]float64, len(w.ColumnWidths))
	for col, width := range w.ColumnWidths {
		if col == idx {
			continue
		}
		if col > idx {
			updated[col-1] = width
			continue
		}
		updated[col] = width
	}
	if len(updated) == 0 {
		w.ColumnWidths = nil
		return
	}
	w.ColumnWidths = updated
}
