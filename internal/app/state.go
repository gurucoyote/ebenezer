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

// SelectionMode represents the current visual selection behavior.
type SelectionMode int

const (
	SelectionNone SelectionMode = iota
	SelectionRange
	SelectionRow
)

// Selection keeps track of anchor/mode for visual selections.
type Selection struct {
	Mode   SelectionMode
	Anchor Cursor
	Active bool
}

// State captures the high-level CLI session data.
type State struct {
	Workbook            *workbook.Workbook
	Cursor              Cursor
	Clipboard           Clipboard
	StyleClipboard      StyleClipboard
	SourcePath          string
	SheetNames          []string
	Selection           Selection
	dirty               bool
	lastSearchQuery     string
	lastSearchForward   bool
	SearchCaseSensitive bool
	csvDelimiter        rune
}

// CSVDelimiter returns the configured delimiter (defaults to workbook.DefaultCSVDelimiter).
func (s *State) CSVDelimiter() rune {
	if s == nil || s.csvDelimiter == 0 {
		return workbook.DefaultCSVDelimiter
	}
	return s.csvDelimiter
}

// IsDirty reports whether the workbook has unsaved changes.
func (s *State) IsDirty() bool {
	if s == nil {
		return false
	}
	return s.dirty
}

func (s *State) markDirty() {
	if s == nil {
		return
	}
	s.dirty = true
}

func (s *State) clearDirty() {
	if s == nil {
		return
	}
	s.dirty = false
}

// SetCSVDelimiter updates the delimiter used for CSV load/save operations.
func (s *State) SetCSVDelimiter(delimiter rune) {
	if s == nil {
		return
	}
	if delimiter == 0 {
		delimiter = workbook.DefaultCSVDelimiter
	}
	s.csvDelimiter = delimiter
}

// ColumnWidthInfo summarizes width metadata for a column.
type ColumnWidthInfo struct {
	Column   int
	Width    float64
	Explicit bool
	Source   string
}

// ClipboardKind describes the type stored in the clipboard.
type ClipboardKind int

const (
	ClipboardNone ClipboardKind = iota
	ClipboardCell
	ClipboardRow
	ClipboardRange
)

// Clipboard stores yank/cut data.
type Clipboard struct {
	Kind        ClipboardKind
	CellValue   string
	Rows        [][]string
	RowStyles   []map[int]workbook.CellStyle
	RangeValues [][]string
	RangeStyles map[int]map[int]workbook.CellStyle
}

type StyleClipboard struct {
	Width  int
	Height int
	Styles []workbook.CellStyle
}

// NewState initializes with sample workbook so the demo has data.
func NewState() *State {
	st := &State{}
	st.csvDelimiter = workbook.DefaultCSVDelimiter
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
	s.Selection = Selection{}
	s.dirty = false
	s.lastSearchQuery = ""
	s.lastSearchForward = true
	s.SearchCaseSensitive = false
	s.Workbook.ActiveCell = activeCell
	if activeCell != "" && s.Goto(activeCell) == nil {
		return
	}
	s.updateActiveCell()
}

// BeginSelection activates visual selection with the current cursor as anchor.
func (s *State) BeginSelection(mode SelectionMode) {
	if s.Workbook == nil {
		return
	}
	s.Selection = Selection{Mode: mode, Anchor: s.Cursor, Active: true}
}

// ToggleSelection switches visual selection on/off depending on the requested mode.
func (s *State) ToggleSelection(mode SelectionMode) {
	if s.Selection.Active && s.Selection.Mode == mode {
		s.ClearSelection()
		return
	}
	s.BeginSelection(mode)
}

// ClearSelection exits visual mode.
func (s *State) ClearSelection() {
	s.Selection = Selection{}
}

// SetSelectionRange activates a rectangular selection based on an A1-style range.
func (s *State) SetSelectionRange(rangeStr string) (string, error) {
	if s.Workbook == nil {
		return "", errors.New("no workbook loaded")
	}
	startRow, startCol, endRow, endCol, err := parseRange(rangeStr)
	if err != nil {
		return "", err
	}
	if startRow < 1 {
		startRow = 1
	}
	if startCol < 1 {
		startCol = 1
	}
	maxRow, maxCol := s.Workbook.MaxCoords()
	if maxRow > 0 && endRow > maxRow {
		endRow = maxRow
	}
	if maxCol > 0 && endCol > maxCol {
		endCol = maxCol
	}
	s.Selection = Selection{
		Mode:   SelectionRange,
		Anchor: Cursor{Row: startRow, Col: startCol},
		Active: true,
	}
	s.Cursor = Cursor{Row: endRow, Col: endCol}
	s.updateActiveCell()
	return s.SelectionSummary(), nil
}

// HasSelection reports whether visual mode is currently active.
func (s *State) HasSelection() bool {
	return s.Selection.Active && s.Selection.Mode != SelectionNone
}

// SelectionBounds returns the inclusive rectangular bounds of the active selection.
func (s *State) SelectionBounds() (startRow, startCol, endRow, endCol int, ok bool) {
	if !s.HasSelection() {
		return 0, 0, 0, 0, false
	}
	startRow, startCol = s.Selection.Anchor.Row, s.Selection.Anchor.Col
	endRow, endCol = s.Cursor.Row, s.Cursor.Col
	if startRow < 1 {
		startRow = 1
	}
	if startCol < 1 {
		startCol = 1
	}
	if endRow < 1 {
		endRow = 1
	}
	if endCol < 1 {
		endCol = 1
	}
	if startRow > endRow {
		startRow, endRow = endRow, startRow
	}
	if startCol > endCol {
		startCol, endCol = endCol, startCol
	}
	if s.Selection.Mode == SelectionRow {
		startCol = 1
		maxRow, maxCol := 0, 0
		if s.Workbook != nil {
			maxRow, maxCol = s.Workbook.MaxCoords()
		}
		if maxRow == 0 {
			maxRow = endRow
		}
		if maxCol == 0 {
			maxCol = 1
		}
		if endRow > maxRow {
			endRow = maxRow
		}
		endCol = maxCol
	}
	return startRow, startCol, endRow, endCol, true
}

// SelectionRowBounds returns only the row portion of the current selection.
func (s *State) SelectionRowBounds() (startRow, endRow int, ok bool) {
	startRow, _, endRow, _, ok = s.SelectionBounds()
	return
}

// SelectionSummary renders a concise description suitable for status output.
func (s *State) SelectionSummary() string {
	if !s.HasSelection() {
		return ""
	}
	startRow, startCol, endRow, endCol, ok := s.SelectionBounds()
	if !ok {
		return ""
	}
	if s.Selection.Mode == SelectionRow {
		if startRow == endRow {
			return fmt.Sprintf("rows %d", startRow)
		}
		return fmt.Sprintf("rows %d-%d", startRow, endRow)
	}
	height := endRow - startRow + 1
	width := endCol - startCol + 1
	return fmt.Sprintf("%s%d:%s%d (%dx%d)", workbook.ColumnName(startCol), startRow, workbook.ColumnName(endCol), endRow, height, width)
}

// LastSearchQuery exposes the most recent search string for prompts.
func (s *State) LastSearchQuery() string {
	return s.lastSearchQuery
}

// Search finds the next cell containing the provided term. When forward is
// false, the search runs in reverse. Matches wrap around the sheet.
func (s *State) Search(term string, forward bool) error {
	if s.Workbook == nil {
		return errors.New("no workbook loaded")
	}
	trimmed := strings.TrimSpace(term)
	if trimmed == "" {
		return errors.New("search term required")
	}
	if !s.performSearch(trimmed, forward) {
		return fmt.Errorf("no match for %q", trimmed)
	}
	s.lastSearchQuery = trimmed
	s.lastSearchForward = forward
	return nil
}

// RepeatSearch reruns the previous search. When sameDirection is false, the
// search direction is flipped (mirroring Vim's `N`).
func (s *State) RepeatSearch(sameDirection bool) error {
	if s.Workbook == nil {
		return errors.New("no workbook loaded")
	}
	if s.lastSearchQuery == "" {
		return errors.New("no previous search")
	}
	forward := s.lastSearchForward
	if !sameDirection {
		forward = !forward
	}
	if !s.performSearch(s.lastSearchQuery, forward) {
		return fmt.Errorf("no match for %q", s.lastSearchQuery)
	}
	s.lastSearchForward = forward
	return nil
}

// SetSearchCaseSensitivity toggles case-sensitive matching.
func (s *State) SetSearchCaseSensitivity(enable bool) {
	s.SearchCaseSensitive = enable
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

// MoveSpan moves the cursor to the next filled cell (or boundary of a filled span)
// in the given direction, similar to Excel's Ctrl+Arrow behavior.
func (s *State) MoveSpan(direction string) error {
	if s.Workbook == nil {
		return errors.New("no workbook loaded")
	}
	maxRow, maxCol := s.Workbook.MaxCoords()
	if maxRow == 0 || maxCol == 0 {
		return errors.New("workbook is empty")
	}

	var deltaRow, deltaCol int
	switch direction {
	case "left":
		deltaCol = -1
	case "right":
		deltaCol = 1
	case "up":
		deltaRow = -1
	case "down":
		deltaRow = 1
	default:
		return fmt.Errorf("unknown direction %s", direction)
	}

	inBounds := func(row, col int) bool {
		return row >= 1 && row <= maxRow && col >= 1 && col <= maxCol
	}
	isFilled := func(row, col int) bool {
		return strings.TrimSpace(s.Workbook.Cell(row, col)) != ""
	}

	row, col := s.Cursor.Row, s.Cursor.Col
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

	curFilled := isFilled(row, col)
	nextRow, nextCol := row+deltaRow, col+deltaCol
	if !inBounds(nextRow, nextCol) {
		s.Cursor = Cursor{Row: row, Col: col}
		s.updateActiveCell()
		return nil
	}

	if curFilled {
		// When already on a filled cell and the next cell is also filled, jump to the
		// edge of the contiguous filled span.
		if isFilled(nextRow, nextCol) {
			for inBounds(nextRow, nextCol) && isFilled(nextRow, nextCol) {
				row, col = nextRow, nextCol
				nextRow, nextCol = row+deltaRow, col+deltaCol
			}
			s.Cursor = Cursor{Row: row, Col: col}
			s.updateActiveCell()
			return nil
		}

		// Otherwise, skip empty cells and land on the next filled cell (or boundary).
		for inBounds(nextRow, nextCol) && !isFilled(nextRow, nextCol) {
			row, col = nextRow, nextCol
			nextRow, nextCol = row+deltaRow, col+deltaCol
		}
		if inBounds(nextRow, nextCol) && isFilled(nextRow, nextCol) {
			row, col = nextRow, nextCol
		}
		s.Cursor = Cursor{Row: row, Col: col}
		s.updateActiveCell()
		return nil
	}

	// Starting from an empty cell: skip empty cells until we find filled data or hit the boundary.
	for inBounds(nextRow, nextCol) && !isFilled(nextRow, nextCol) {
		row, col = nextRow, nextCol
		nextRow, nextCol = row+deltaRow, col+deltaCol
	}
	if inBounds(nextRow, nextCol) && isFilled(nextRow, nextCol) {
		row, col = nextRow, nextCol
	}
	s.Cursor = Cursor{Row: row, Col: col}
	s.updateActiveCell()
	return nil
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
	if err := s.Workbook.Save(path, workbook.WithCSVDelimiter(s.CSVDelimiter())); err != nil {
		return err
	}
	s.SourcePath = path
	s.Workbook.Name = path
	if len(s.SheetNames) == 0 {
		s.SheetNames = []string{s.Workbook.Sheet}
	}
	s.clearDirty()
	return nil
}

// EditCurrentCell sets the current cell to the provided value.
func (s *State) EditCurrentCell(value string) {
	if s.Workbook == nil {
		return
	}
	s.Workbook.SetCell(s.Cursor.Row, s.Cursor.Col, value)
	s.markDirty()
	s.updateActiveCell()
}

// ClearCurrentCell blanks the current cell.
func (s *State) ClearCurrentCell() {
	if s.Workbook == nil {
		return
	}
	s.markDirty()
	if s.HasSelection() {
		if s.Selection.Mode == SelectionRow {
			if startRow, endRow, ok := s.SelectionRowBounds(); ok {
				s.deleteRowRange(startRow, endRow)
			}
		} else {
			if startRow, startCol, endRow, endCol, ok := s.SelectionBounds(); ok {
				s.clearRange(startRow, startCol, endRow, endCol)
			}
		}
		s.ClearSelection()
		s.updateActiveCell()
		return
	}
	s.Workbook.ClearCell(s.Cursor.Row, s.Cursor.Col)
	s.updateActiveCell()
}

// YankCurrentCell copies the current cell into the clipboard without mutation.
func (s *State) YankCurrentCell() string {
	if s.Workbook == nil {
		return ""
	}
	if s.HasSelection() {
		if s.Selection.Mode == SelectionRow {
			if startRow, endRow, ok := s.SelectionRowBounds(); ok {
				rows, styles := s.collectRows(startRow, endRow)
				s.Clipboard = Clipboard{Kind: ClipboardRow, Rows: rows, RowStyles: styles}
				value := ""
				if len(rows) > 0 && len(rows[0]) > 0 {
					value = rows[0][0]
				}
				s.ClearSelection()
				return value
			}
		}
		if startRow, startCol, endRow, endCol, ok := s.SelectionBounds(); ok {
			values, styles := s.collectRangeValues(startRow, startCol, endRow, endCol)
			s.Clipboard = Clipboard{Kind: ClipboardRange, RangeValues: values, RangeStyles: styles}
			value := ""
			if len(values) > 0 && len(values[0]) > 0 {
				value = values[0][0]
			}
			s.ClearSelection()
			return value
		}
		s.ClearSelection()
		return ""
	}
	value := s.CurrentValue()
	s.Clipboard = Clipboard{Kind: ClipboardCell, CellValue: value}
	return value
}

// CutCurrentCell copies the current cell into the clipboard and clears it.
func (s *State) CutCurrentCell() string {
	if s.Workbook == nil {
		return ""
	}
	s.markDirty()
	if s.HasSelection() {
		if s.Selection.Mode == SelectionRow {
			if startRow, endRow, ok := s.SelectionRowBounds(); ok {
				rows, styles := s.collectRows(startRow, endRow)
				s.Clipboard = Clipboard{Kind: ClipboardRow, Rows: rows, RowStyles: styles}
				s.deleteRowRange(startRow, endRow)
				s.ClearSelection()
				s.updateActiveCell()
				if len(rows) > 0 && len(rows[0]) > 0 {
					return rows[0][0]
				}
				return ""
			}
		}
		if startRow, startCol, endRow, endCol, ok := s.SelectionBounds(); ok {
			values, styles := s.collectRangeValues(startRow, startCol, endRow, endCol)
			s.Clipboard = Clipboard{Kind: ClipboardRange, RangeValues: values, RangeStyles: styles}
			s.clearRange(startRow, startCol, endRow, endCol)
			s.ClearSelection()
			s.updateActiveCell()
			if len(values) > 0 && len(values[0]) > 0 {
				return values[0][0]
			}
			return ""
		}
		s.ClearSelection()
		return ""
	}
	value := s.CurrentValue()
	s.Clipboard = Clipboard{Kind: ClipboardCell, CellValue: value}
	s.Workbook.ClearCell(s.Cursor.Row, s.Cursor.Col)
	s.updateActiveCell()
	return value
}

// PasteClipboard writes the clipboard contents into the sheet. When before is
// true, row pastes happen above the current row; otherwise below.
func (s *State) PasteClipboard(before bool) error {
	if s.Workbook == nil {
		return errors.New("no workbook loaded")
	}
	targetRow := s.Cursor.Row
	targetCol := s.Cursor.Col
	targetBefore := before
	destHeight, destWidth := 0, 0
	if s.HasSelection() {
		if s.Selection.Mode == SelectionRow {
			if startRow, endRow, ok := s.SelectionRowBounds(); ok {
				s.deleteRowRange(startRow, endRow)
				targetRow = startRow
				s.Cursor.Row = startRow
				targetBefore = true
			}
		} else if startRow, startCol, endRow, endCol, ok := s.SelectionBounds(); ok {
			targetRow = startRow
			targetCol = startCol
			destHeight = endRow - startRow + 1
			destWidth = endCol - startCol + 1
			s.Cursor = Cursor{Row: startRow, Col: startCol}
		}
		s.ClearSelection()
	}
	switch s.Clipboard.Kind {
	case ClipboardCell:
		if destHeight > 0 && destWidth > 0 {
			for r := 0; r < destHeight; r++ {
				for c := 0; c < destWidth; c++ {
					s.Workbook.SetCell(targetRow+r, targetCol+c, s.Clipboard.CellValue)
				}
			}
			s.markDirty()
			return nil
		}
		s.Workbook.SetCell(targetRow, targetCol, s.Clipboard.CellValue)
		s.markDirty()
		return nil
	case ClipboardRow:
		if len(s.Clipboard.Rows) == 0 {
			return errors.New("clipboard empty")
		}
		target := targetRow
		if !targetBefore {
			target++
		}
		for i, rowValues := range s.Clipboard.Rows {
			insertIdx := target + i
			row := make([]string, len(rowValues))
			copy(row, rowValues)
			s.Workbook.InsertRow(insertIdx, row)
			if i < len(s.Clipboard.RowStyles) {
				for col, style := range s.Clipboard.RowStyles[i] {
					s.Workbook.SetStyle(insertIdx, col, style)
				}
			}
		}
		s.markDirty()
		return nil
	case ClipboardRange:
		if len(s.Clipboard.RangeValues) == 0 {
			return errors.New("clipboard empty")
		}
		clipHeight := len(s.Clipboard.RangeValues)
		clipWidth := 0
		if clipHeight > 0 {
			clipWidth = len(s.Clipboard.RangeValues[0])
		}
		if clipWidth == 0 {
			return errors.New("clipboard empty")
		}
		if destHeight > 0 && destWidth > 0 {
			if clipHeight == 1 && clipWidth == 1 {
				value := s.Clipboard.RangeValues[0][0]
				var style workbook.CellStyle
				if rowStyle, ok := s.Clipboard.RangeStyles[0]; ok {
					style = rowStyle[0]
				}
				for r := 0; r < destHeight; r++ {
					for c := 0; c < destWidth; c++ {
						s.Workbook.SetCell(targetRow+r, targetCol+c, value)
						if !style.Empty() {
							s.Workbook.SetStyle(targetRow+r, targetCol+c, style)
						}
					}
				}
				s.markDirty()
				return nil
			}
			if clipHeight != destHeight || clipWidth != destWidth {
				return errors.New("destination selection size must match copied range")
			}
			if err := s.applyRangeClipboard(targetRow, targetCol, clipHeight, clipWidth); err != nil {
				return err
			}
			s.markDirty()
			return nil
		}
		if err := s.applyRangeClipboard(targetRow, targetCol, clipHeight, clipWidth); err != nil {
			return err
		}
		s.markDirty()
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
	if row == nil {
		row = []string{}
	}
	s.Clipboard = Clipboard{Kind: ClipboardRow, Rows: [][]string{row}, RowStyles: []map[int]workbook.CellStyle{styles}}
	return row
}

// CutCurrentRow copies the row to the clipboard and removes it from the sheet.
func (s *State) CutCurrentRow() []string {
	row := s.YankCurrentRow()
	if row == nil {
		return nil
	}
	s.markDirty()
	s.deleteRow(s.Cursor.Row)
	return row
}

// DeleteCurrentRow removes the row without touching the clipboard.
func (s *State) DeleteCurrentRow() {
	if s.Workbook == nil {
		return
	}
	s.markDirty()
	s.deleteRow(s.Cursor.Row)
}

// InsertColumnLeft inserts a blank column before the cursor.
func (s *State) InsertColumnLeft() {
	if s.Workbook == nil {
		return
	}
	s.Workbook.InsertColumn(s.Cursor.Col)
	s.markDirty()
	s.updateActiveCell()
}

// InsertColumnRight inserts a blank column after the cursor.
func (s *State) InsertColumnRight() {
	if s.Workbook == nil {
		return
	}
	s.Workbook.InsertColumn(s.Cursor.Col + 1)
	s.markDirty()
	s.updateActiveCell()
}

// DeleteCurrentColumn removes the column at the cursor.
func (s *State) DeleteCurrentColumn() {
	if s.Workbook == nil {
		return
	}
	if _, ok := s.Workbook.DeleteColumn(s.Cursor.Col); !ok {
		return
	}
	s.markDirty()
	_, maxCol := s.Workbook.MaxCoords()
	if maxCol == 0 {
		s.Cursor.Col = 1
	} else if s.Cursor.Col > maxCol {
		s.Cursor.Col = maxCol
	}
	s.updateActiveCell()
}

// ResolveColumnSpan converts a user-provided column token (e.g., "A:D",
// "3", "B2:G5") into inclusive column indexes, defaulting to the full sheet
// when empty.
func (s *State) ResolveColumnSpan(spec string) (int, int, error) {
	if s.Workbook == nil {
		return 0, 0, errors.New("no workbook loaded")
	}
	_, maxCol := s.Workbook.MaxCoords()
	if maxCol == 0 {
		maxCol = 1
	}
	trimmed := strings.TrimSpace(spec)
	if trimmed == "" {
		return 1, maxCol, nil
	}
	start, end, err := workbook.ParseColumnSpec(trimmed)
	if err != nil {
		return 0, 0, err
	}
	if start < 1 {
		start = 1
	}
	if end < 1 {
		end = 1
	}
	if start > maxCol {
		start = maxCol
	}
	if end > maxCol {
		end = maxCol
	}
	if start > end {
		start, end = end, start
	}
	return start, end, nil
}

// ColumnWidths returns width metadata for a span of columns.
func (s *State) ColumnWidths(startCol, endCol int) ([]ColumnWidthInfo, error) {
	if s.Workbook == nil {
		return nil, errors.New("no workbook loaded")
	}
	if startCol < 1 || endCol < 1 {
		return nil, fmt.Errorf("column indexes must be >= 1")
	}
	if startCol > endCol {
		startCol, endCol = endCol, startCol
	}
	infos := make([]ColumnWidthInfo, 0, endCol-startCol+1)
	for col := startCol; col <= endCol; col++ {
		width, ok := s.Workbook.ColumnWidth(col)
		if !ok {
			width = workbook.DefaultColumnWidth
		}
		infos = append(infos, ColumnWidthInfo{
			Column:   col,
			Width:    width,
			Explicit: ok,
			Source:   "current",
		})
	}
	return infos, nil
}

// SetColumnWidth sets an explicit width (in Excel units) across the span.
func (s *State) SetColumnWidth(startCol, endCol int, width float64) ([]ColumnWidthInfo, error) {
	if s.Workbook == nil {
		return nil, errors.New("no workbook loaded")
	}
	if width <= 0 {
		return nil, fmt.Errorf("width must be greater than zero")
	}
	if startCol > endCol {
		startCol, endCol = endCol, startCol
	}
	infos := make([]ColumnWidthInfo, 0, endCol-startCol+1)
	for col := startCol; col <= endCol; col++ {
		s.Workbook.SetColumnWidth(col, width)
		infos = append(infos, ColumnWidthInfo{
			Column:   col,
			Width:    width,
			Explicit: true,
			Source:   "set",
		})
	}
	s.markDirty()
	return infos, nil
}

// AutoColumnWidth estimates and applies widths across the span using the
// provided heuristic options.
func (s *State) AutoColumnWidth(startCol, endCol int, opts workbook.ColumnWidthOptions) ([]ColumnWidthInfo, error) {
	if s.Workbook == nil {
		return nil, errors.New("no workbook loaded")
	}
	if startCol > endCol {
		startCol, endCol = endCol, startCol
	}
	infos := make([]ColumnWidthInfo, 0, endCol-startCol+1)
	for col := startCol; col <= endCol; col++ {
		width := s.Workbook.EstimateColumnWidth(col, opts)
		s.Workbook.SetColumnWidth(col, width)
		infos = append(infos, ColumnWidthInfo{
			Column:   col,
			Width:    width,
			Explicit: true,
			Source:   "auto",
		})
	}
	s.markDirty()
	return infos, nil
}

func (s *State) deleteRow(idx int) {
	if s.Workbook == nil {
		return
	}
	if _, ok := s.Workbook.DeleteRow(idx); !ok {
		return
	}
	s.markDirty()
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
	s.markDirty()
	s.updateActiveCell()
}

// InsertRowBelow inserts a blank row after the cursor.
func (s *State) InsertRowBelow() {
	if s.Workbook == nil {
		return
	}
	s.Workbook.InsertRow(s.Cursor.Row+1, nil)
	s.markDirty()
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

func (s *State) collectRows(startRow, endRow int) ([][]string, []map[int]workbook.CellStyle) {
	if s.Workbook == nil {
		return nil, nil
	}
	if startRow > endRow {
		startRow, endRow = endRow, startRow
	}
	count := endRow - startRow + 1
	if count < 1 {
		return nil, nil
	}
	rows := make([][]string, 0, count)
	styles := make([]map[int]workbook.CellStyle, 0, count)
	for row := startRow; row <= endRow; row++ {
		data := s.Workbook.Row(row)
		if data == nil {
			data = []string{}
		}
		rows = append(rows, data)
		styles = append(styles, s.collectRowStyles(row))
	}
	return rows, styles
}

func (s *State) collectRangeValues(startRow, startCol, endRow, endCol int) ([][]string, map[int]map[int]workbook.CellStyle) {
	values := [][]string{}
	styles := map[int]map[int]workbook.CellStyle{}
	if s.Workbook == nil {
		return values, styles
	}
	if startRow > endRow {
		startRow, endRow = endRow, startRow
	}
	if startCol > endCol {
		startCol, endCol = endCol, startCol
	}
	height := endRow - startRow + 1
	width := endCol - startCol + 1
	values = make([][]string, height)
	for r := 0; r < height; r++ {
		values[r] = make([]string, width)
		absRow := startRow + r
		for c := 0; c < width; c++ {
			absCol := startCol + c
			values[r][c] = s.Workbook.Cell(absRow, absCol)
			addr := fmt.Sprintf("%s%d", workbook.ColumnName(absCol), absRow)
			if style, ok := s.Workbook.Style(addr); ok {
				if styles[r] == nil {
					styles[r] = map[int]workbook.CellStyle{}
				}
				styles[r][c] = style
			}
		}
	}
	return values, styles
}

// ExportRange returns the values for the given range (or current cell when empty).
func (s *State) ExportRange(rangeStr string) ([][]string, error) {
	values, _, _, err := s.ExportRangeWithStyles(rangeStr)
	return values, err
}

// ExportRangeWithStyles returns values, style metadata, and the normalized range string.
func (s *State) ExportRangeWithStyles(rangeStr string) ([][]string, map[int]map[int]workbook.CellStyle, string, error) {
	startRow, startCol, endRow, endCol, normalized, err := s.resolveRange(rangeStr)
	if err != nil {
		return nil, nil, "", err
	}
	values, styles := s.collectRangeValues(startRow, startCol, endRow, endCol)
	if len(values) == 0 || len(values[0]) == 0 {
		return nil, nil, "", errors.New("invalid range dimensions")
	}
	return values, styles, normalized, nil
}

func (s *State) clearRange(startRow, startCol, endRow, endCol int) {
	if s.Workbook == nil {
		return
	}
	if startRow > endRow {
		startRow, endRow = endRow, startRow
	}
	if startCol > endCol {
		startCol, endCol = endCol, startCol
	}
	for row := startRow; row <= endRow; row++ {
		for col := startCol; col <= endCol; col++ {
			s.Workbook.ClearCell(row, col)
		}
	}
}

func (s *State) deleteRowRange(start, end int) {
	if s.Workbook == nil {
		return
	}
	if start > end {
		start, end = end, start
	}
	for row := start; row <= end; row++ {
		s.deleteRow(start)
	}
	maxRow, _ := s.Workbook.MaxCoords()
	if maxRow == 0 {
		s.Cursor.Row = 1
		return
	}
	if start > maxRow {
		s.Cursor.Row = maxRow
		return
	}
	s.Cursor.Row = start
}

func (s *State) applyRangeClipboard(rowStart, colStart, height, width int) error {
	for r := 0; r < height && r < len(s.Clipboard.RangeValues); r++ {
		rowValues := s.Clipboard.RangeValues[r]
		for c := 0; c < width && c < len(rowValues); c++ {
			value := rowValues[c]
			row := rowStart + r
			col := colStart + c
			s.Workbook.SetCell(row, col, value)
			if rowStyle, ok := s.Clipboard.RangeStyles[r]; ok {
				if style, ok := rowStyle[c]; ok {
					s.Workbook.SetStyle(row, col, style)
				}
			}
		}
	}
	return nil
}

func (s *State) performSearch(term string, forward bool) bool {
	if s.Workbook == nil {
		return false
	}
	maxRow, maxCol := s.Workbook.MaxCoords()
	if maxRow == 0 || maxCol == 0 {
		return false
	}
	target := term
	caseInsensitive := !s.SearchCaseSensitive
	if caseInsensitive {
		target = strings.ToLower(term)
	}
	row, col := s.Cursor.Row, s.Cursor.Col
	total := maxRow * maxCol
	for steps := 0; steps < total; steps++ {
		row, col = advancePosition(row, col, forward, maxRow, maxCol)
		value := s.Workbook.Cell(row, col)
		if caseInsensitive {
			value = strings.ToLower(value)
		}
		if strings.Contains(value, target) {
			s.Cursor = Cursor{Row: row, Col: col}
			s.ClearSelection()
			s.updateActiveCell()
			return true
		}
	}
	return false
}

func advancePosition(row, col int, forward bool, maxRow, maxCol int) (int, int) {
	if forward {
		col++
		if col > maxCol {
			col = 1
			row++
			if row > maxRow {
				row = 1
			}
		}
		return row, col
	}
	col--
	if col < 1 {
		col = maxCol
		row--
		if row < 1 {
			row = maxRow
		}
	}
	return row, col
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
	s.markDirty()
	return nil
}

// ApplyStyles applies a payload of styles to the provided range. When the
// payload is a single cell, it is broadcast across the destination. Otherwise
// the payload dimensions must match the destination exactly.
func (s *State) ApplyStyles(rangeStr string, styles [][]workbook.CellStyle) error {
	if s.Workbook == nil {
		return errors.New("no workbook loaded")
	}
	if len(styles) == 0 || len(styles[0]) == 0 {
		return errors.New("styles payload required")
	}
	startRow, startCol, endRow, endCol, _, err := s.resolveRange(rangeStr)
	if err != nil {
		return err
	}
	height := endRow - startRow + 1
	width := endCol - startCol + 1
	payloadHeight := len(styles)
	payloadWidth := len(styles[0])
	single := payloadHeight == 1 && payloadWidth == 1
	for _, row := range styles {
		if len(row) != payloadWidth {
			return errors.New("styles payload rows must be equal length")
		}
	}
	if !single && (payloadHeight != height || payloadWidth != width) {
		return errors.New("styles payload dimensions must match destination range")
	}
	for r := 0; r < height; r++ {
		for c := 0; c < width; c++ {
			row := startRow + r
			col := startCol + c
			var style workbook.CellStyle
			if single {
				style = styles[0][0]
			} else {
				style = styles[r][c]
			}
			s.Workbook.SetStyle(row, col, style)
		}
	}
	s.markDirty()
	return nil
}

func parseRangeOrDefault(rangeStr, fallback string) (int, int, int, int, error) {
	input := strings.TrimSpace(rangeStr)
	if input == "" {
		input = fallback
	}
	return parseRange(input)
}

func (s *State) resolveRange(rangeStr string) (int, int, int, int, string, error) {
	if s.Workbook == nil {
		return 0, 0, 0, 0, "", errors.New("no workbook loaded")
	}
	startRow, startCol, endRow, endCol, err := parseRangeOrDefault(rangeStr, s.Address())
	if err != nil {
		return 0, 0, 0, 0, "", err
	}
	return startRow, startCol, endRow, endCol, formatRangeString(startRow, startCol, endRow, endCol), nil
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

func formatRangeString(startRow, startCol, endRow, endCol int) string {
	if startRow == endRow && startCol == endCol {
		return fmt.Sprintf("%s%d", workbook.ColumnName(startCol), startRow)
	}
	return fmt.Sprintf("%s%d:%s%d", workbook.ColumnName(startCol), startRow, workbook.ColumnName(endCol), endRow)
}
