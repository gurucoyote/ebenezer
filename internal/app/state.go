package app

import (
	"fmt"

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

// Address returns Excel-like cell reference (e.g., A1).
func (s *State) Address() string {
	return fmt.Sprintf("%s%d", columnName(s.Cursor.Col), s.Cursor.Row)
}

func columnName(col int) string {
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
