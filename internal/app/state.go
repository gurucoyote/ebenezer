package app

import (
	"errors"
	"fmt"
	"strconv"
	"strings"
	"unicode"

	"ebenezer/internal/workbook"
)

// Cursor tracks 1-based spreadsheet coordinates.
type Cursor struct {
	Row int
	Col int
}

// State captures the high-level CLI session data.
type State struct {
	Workbook       *workbook.Workbook
	Cursor         Cursor
	Clipboard      Clipboard
	StyleClipboard StyleClipboard
	SourcePath     string
	SheetNames     []string
}

// ClipboardKind describes the type stored in the clipboard.
type ClipboardKind int

const (
	ClipboardNone ClipboardKind = iota
	ClipboardCell
	ClipboardRow
)

// Clipboard stores yank/cut data.
type Clipboard struct {
	Kind      ClipboardKind
	CellValue string
	RowValues []string
	RowStyles map[int]workbook.CellStyle
}

type StyleClipboard struct {
	Width  int
	Height int
	Styles []workbook.CellStyle
}

// NewState initializes with sample workbook so the demo has data.
func NewState() *State {
	st := &State{}
	wb := workbook.SampleWorkbook()
	st.LoadWorkbook(wb, "", []string{wb.Sheet}, "")
	return st
}

// LoadWorkbook swaps the active workbook and resets the cursor.
func (s *State) LoadWorkbook(wb *workbook.Workbook, path string, sheets []string, activeCell string) {
	if wb == nil {
		return
	}
	s.Workbook = wb
	s.SourcePath = path
	s.SheetNames = append([]string(nil), sheets...)
	s.Cursor = Cursor{Row: 1, Col: 1}
	s.Clipboard = Clipboard{}
	s.StyleClipboard = StyleClipboard{}
	s.Workbook.ActiveCell = activeCell
	if activeCell != "" && s.Goto(activeCell) == nil {
		return
	}
	s.updateActiveCell()
}

// Move adjusts the cursor, clamping to valid coordinates.
func (s *State) Move(deltaRow, deltaCol int) {
	if s.Workbook == nil {
		return
	}
	s.Cursor.Row += deltaRow
	s.Cursor.Col += deltaCol

	if s.Cursor.Row < 1 {
		s.Cursor.Row = 1
	}
	if s.Cursor.Col < 1 {
		s.Cursor.Col = 1
	}

	maxRow, maxCol := s.Workbook.MaxCoords()
	if maxRow > 0 && s.Cursor.Row > maxRow {
		s.Cursor.Row = maxRow
	}
	if maxCol > 0 && s.Cursor.Col > maxCol {
		s.Cursor.Col = maxCol
	}
	s.updateActiveCell()
}

// CurrentValue returns the cell value at the cursor.
func (s *State) CurrentValue() string {
	if s.Workbook == nil {
		return ""
	}
	return s.Workbook.Cell(s.Cursor.Row, s.Cursor.Col)
}

// ColumnHeader returns row 1 of the current column.
func (s *State) ColumnHeader() string {
	if s.Workbook == nil {
		return ""
	}
	return s.Workbook.Cell(1, s.Cursor.Col)
}

// RowHeader returns column 1 of the current row.
func (s *State) RowHeader() string {
	if s.Workbook == nil {
		return ""
	}
	return s.Workbook.Cell(s.Cursor.Row, 1)
}

// Goto moves the cursor to the provided address (e.g., B12).
func (s *State) Goto(address string) error {
	if s.Workbook == nil {
		return errors.New("no workbook loaded")
	}
	row, col, err := parseAddress(address)
	if err != nil {
		return err
	}
	maxRow, maxCol := s.Workbook.MaxCoords()
	if maxRow == 0 || maxCol == 0 {
		return errors.New("workbook is empty")
	}
	if row < 1 {
		row = 1
	}
	if col < 1 {
		col = 1
	}
	if row > maxRow {
		row = maxRow
	}
	if col > maxCol {
		col = maxCol
	}
	s.Cursor = Cursor{Row: row, Col: col}
	s.updateActiveCell()
	return nil
}

// StyleAt returns style metadata for the provided cell address, or the
// current cursor when address is empty.
func (s *State) StyleAt(address string) (workbook.CellStyle, bool) {
	if s.Workbook == nil {
		return workbook.CellStyle{}, false
	}
	if strings.TrimSpace(address) == "" {
		address = s.Address()
	}
	return s.Workbook.Style(address)
}

// Save writes the workbook to disk.
func (s *State) Save(path string) error {
	if s.Workbook == nil {
		return errors.New("no workbook loaded")
	}
	if strings.TrimSpace(path) == "" {
		return errors.New("filename required")
	}
	s.updateActiveCell()
	if err := s.Workbook.Save(path); err != nil {
		return err
	}
	s.SourcePath = path
	s.Workbook.Name = path
	if len(s.SheetNames) == 0 {
		s.SheetNames = []string{s.Workbook.Sheet}
	}
	return nil
}

// EditCurrentCell sets the current cell to the provided value.
func (s *State) EditCurrentCell(value string) {
	if s.Workbook == nil {
		return
	}
	s.Workbook.SetCell(s.Cursor.Row, s.Cursor.Col, value)
	s.updateActiveCell()
}

// ClearCurrentCell blanks the current cell.
func (s *State) ClearCurrentCell() {
	if s.Workbook == nil {
		return
	}
	s.Workbook.ClearCell(s.Cursor.Row, s.Cursor.Col)
	s.updateActiveCell()
}

// YankCurrentCell copies the current cell into the clipboard without mutation.
func (s *State) YankCurrentCell() string {
	value := s.CurrentValue()
	s.Clipboard = Clipboard{Kind: ClipboardCell, CellValue: value}
	return value
}

// CutCurrentCell copies the current cell into the clipboard and clears it.
func (s *State) CutCurrentCell() string {
	value := s.CurrentValue()
	s.Clipboard = Clipboard{Kind: ClipboardCell, CellValue: value}
	s.ClearCurrentCell()
	return value
}

// PasteClipboard writes the clipboard contents into the sheet. When before is
// true, row pastes happen above the current row; otherwise below.
func (s *State) PasteClipboard(before bool) error {
	if s.Workbook == nil {
		return errors.New("no workbook loaded")
	}
	switch s.Clipboard.Kind {
	case ClipboardCell:
		s.Workbook.SetCell(s.Cursor.Row, s.Cursor.Col, s.Clipboard.CellValue)
		return nil
	case ClipboardRow:
		row := make([]string, len(s.Clipboard.RowValues))
		copy(row, s.Clipboard.RowValues)
		target := s.Cursor.Row
		if !before {
			target++
		}
		s.Workbook.InsertRow(target, row)
		for col, style := range s.Clipboard.RowStyles {
			s.Workbook.SetStyle(target, col, style)
		}
		return nil
	default:
		return errors.New("clipboard empty")
	}
}

// YankCurrentRow copies the entire row into the clipboard.
func (s *State) YankCurrentRow() []string {
	if s.Workbook == nil {
		return nil
	}
	row := s.Workbook.Row(s.Cursor.Row)
	styles := s.collectRowStyles(s.Cursor.Row)
	s.Clipboard = Clipboard{Kind: ClipboardRow, RowValues: row, RowStyles: styles}
	return row
}

// CutCurrentRow copies the row to the clipboard and removes it from the sheet.
func (s *State) CutCurrentRow() []string {
	row := s.YankCurrentRow()
	if row == nil {
		return nil
	}
	s.deleteRow(s.Cursor.Row)
	return row
}

// DeleteCurrentRow removes the row without touching the clipboard.
func (s *State) DeleteCurrentRow() {
	if s.Workbook == nil {
		return
	}
	s.deleteRow(s.Cursor.Row)
}

func (s *State) deleteRow(idx int) {
	if s.Workbook == nil {
		return
	}
	if _, ok := s.Workbook.DeleteRow(idx); !ok {
		return
	}
	maxRow, _ := s.Workbook.MaxCoords()
	if maxRow == 0 {
		s.Cursor.Row = 1
	} else if s.Cursor.Row > maxRow {
		s.Cursor.Row = maxRow
	}
	s.updateActiveCell()
}

// InsertRowAbove inserts a blank row before the cursor.
func (s *State) InsertRowAbove() {
	if s.Workbook == nil {
		return
	}
	s.Workbook.InsertRow(s.Cursor.Row, nil)
	s.updateActiveCell()
}

// InsertRowBelow inserts a blank row after the cursor.
func (s *State) InsertRowBelow() {
	if s.Workbook == nil {
		return
	}
	s.Workbook.InsertRow(s.Cursor.Row+1, nil)
	s.updateActiveCell()
}

func (s *State) collectRowStyles(row int) map[int]workbook.CellStyle {
	styles := map[int]workbook.CellStyle{}
	if s.Workbook == nil {
		return styles
	}
	_, maxCol := s.Workbook.MaxCoords()
	for col := 1; col <= maxCol; col++ {
		addr := fmt.Sprintf("%s%d", workbook.ColumnName(col), row)
		if style, ok := s.Workbook.Style(addr); ok {
			styles[col] = style
		}
	}
	return styles
}

// Address returns Excel-like cell reference (e.g., A1).
func (s *State) Address() string {
	return fmt.Sprintf("%s%d", workbook.ColumnName(s.Cursor.Col), s.Cursor.Row)
}

func parseAddress(address string) (int, int, error) {
	addr := strings.TrimSpace(address)
	if addr == "" {
		return 0, 0, errors.New("address required")
	}
	addr = strings.ToUpper(addr)
	var letters, digits strings.Builder
	for _, r := range addr {
		switch {
		case unicode.IsLetter(r):
			letters.WriteRune(r)
		case unicode.IsDigit(r):
			digits.WriteRune(r)
		default:
			return 0, 0, fmt.Errorf("invalid character %q in address", r)
		}
	}
	if letters.Len() == 0 || digits.Len() == 0 {
		return 0, 0, errors.New("address must include column letters and row digits")
	}
	col := lettersToNumber(letters.String())
	row, err := strconv.Atoi(digits.String())
	if err != nil {
		return 0, 0, fmt.Errorf("invalid row digits: %w", err)
	}
	return row, col, nil
}

func lettersToNumber(s string) int {
	result := 0
	for _, r := range s {
		result = result*26 + int(r-'A'+1)
	}
	return result
}

func (s *State) updateActiveCell() {
	if s.Workbook == nil {
		return
	}
	s.Workbook.ActiveCell = s.Address()
}

// CopyStyle copies formatting from the provided range (defaults to current cell).
func (s *State) CopyStyle(rangeStr string) error {
	if s.Workbook == nil {
		return errors.New("no workbook loaded")
	}
	startRow, startCol, endRow, endCol, err := parseRangeOrDefault(rangeStr, s.Address())
	if err != nil {
		return err
	}
	var styles []workbook.CellStyle
	for r := startRow; r <= endRow; r++ {
		for c := startCol; c <= endCol; c++ {
			addr := fmt.Sprintf("%s%d", workbook.ColumnName(c), r)
			style, ok := s.Workbook.Style(addr)
			if !ok {
				style = workbook.CellStyle{}
			}
			styles = append(styles, style)
		}
	}
	s.StyleClipboard = StyleClipboard{
		Width:  endCol - startCol + 1,
		Height: endRow - startRow + 1,
		Styles: styles,
	}
	return nil
}

// PasteStyle applies copied formatting to the provided range (defaults to current cell).
func (s *State) PasteStyle(rangeStr string) error {
	if s.Workbook == nil {
		return errors.New("no workbook loaded")
	}
	if len(s.StyleClipboard.Styles) == 0 {
		return errors.New("style clipboard empty")
	}
	startRow, startCol, endRow, endCol, err := parseRangeOrDefault(rangeStr, s.Address())
	if err != nil {
		return err
	}
	destWidth := endCol - startCol + 1
	destHeight := endRow - startRow + 1

	switch {
	case s.StyleClipboard.Width == 1 && s.StyleClipboard.Height == 1:
		style := s.StyleClipboard.Styles[0]
		for r := startRow; r <= endRow; r++ {
			for c := startCol; c <= endCol; c++ {
				s.Workbook.SetStyle(r, c, style)
			}
		}
	case s.StyleClipboard.Width == destWidth && s.StyleClipboard.Height == destHeight:
		idx := 0
		for r := startRow; r <= endRow; r++ {
			for c := startCol; c <= endCol; c++ {
				s.Workbook.SetStyle(r, c, s.StyleClipboard.Styles[idx])
				idx++
			}
		}
	default:
		return errors.New("destination range size must match copied style range")
	}
	return nil
}

func parseRangeOrDefault(rangeStr, fallback string) (int, int, int, int, error) {
	input := strings.TrimSpace(rangeStr)
	if input == "" {
		input = fallback
	}
	return parseRange(input)
}

func parseRange(rangeStr string) (int, int, int, int, error) {
	parts := strings.Split(rangeStr, ":")
	if len(parts) > 2 {
		return 0, 0, 0, 0, fmt.Errorf("invalid range %s", rangeStr)
	}
	startAddr := strings.TrimSpace(parts[0])
	startRow, startCol, err := parseAddress(startAddr)
	if err != nil {
		return 0, 0, 0, 0, err
	}
	endRow, endCol := startRow, startCol
	if len(parts) == 2 {
		endAddr := strings.TrimSpace(parts[1])
		endRow, endCol, err = parseAddress(endAddr)
		if err != nil {
			return 0, 0, 0, 0, err
		}
	}
	if startRow > endRow {
		startRow, endRow = endRow, startRow
	}
	if startCol > endCol {
		startCol, endCol = endCol, startCol
	}
	return startRow, startCol, endRow, endCol, nil
}
