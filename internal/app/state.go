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
	Workbook *workbook.Workbook
	Cursor   Cursor
}

// NewState initializes with sample workbook so the demo has data.
func NewState() *State {
	wb := workbook.SampleWorkbook()
	return &State{
		Workbook: wb,
		Cursor:   Cursor{Row: 1, Col: 1},
	}
}

// LoadWorkbook swaps the active workbook and resets the cursor.
func (s *State) LoadWorkbook(wb *workbook.Workbook) {
	if wb == nil {
		return
	}
	s.Workbook = wb
	s.Cursor = Cursor{Row: 1, Col: 1}
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
