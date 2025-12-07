package app

import (
	"testing"

	"ebenezer/internal/workbook"
)

func TestMoveClampsWithinBounds(t *testing.T) {
	st := &State{Workbook: workbook.SampleWorkbook(), Cursor: Cursor{Row: 1, Col: 1}}

	st.Move(-10, -10)
	if st.Cursor.Row != 1 || st.Cursor.Col != 1 {
		t.Fatalf("expected cursor to clamp at 1,1 got %d,%d", st.Cursor.Row, st.Cursor.Col)
	}

	st.Move(999, 999)
	maxRow, maxCol := st.Workbook.MaxCoords()
	if st.Cursor.Row != maxRow || st.Cursor.Col != maxCol {
		t.Fatalf("expected cursor to clamp at %d,%d got %d,%d", maxRow, maxCol, st.Cursor.Row, st.Cursor.Col)
	}
}

func TestAddressAndValue(t *testing.T) {
	st := NewState()
	if got, want := st.Address(), "A1"; got != want {
		t.Fatalf("expected address %s, got %s", want, got)
	}
	if got, want := st.CurrentValue(), "Item"; got != want {
		t.Fatalf("expected value %s, got %s", want, got)
	}

	st.Move(2, 1) // now at row3 col2
	if got, want := st.Address(), "B3"; got != want {
		t.Fatalf("expected address %s, got %s", want, got)
	}
	if got, want := st.CurrentValue(), "5"; got != want {
		t.Fatalf("expected value %s, got %s", want, got)
	}
}

func TestGotoAndHeaders(t *testing.T) {
	st := NewState()
	if err := st.Goto("B2"); err != nil {
		t.Fatalf("goto failed: %v", err)
	}
	if got, want := st.Address(), "B2"; got != want {
		t.Fatalf("expected address %s, got %s", want, got)
	}
	if got, want := st.ColumnHeader(), "Qty"; got != want {
		t.Fatalf("expected column header %s, got %s", want, got)
	}
	if got, want := st.RowHeader(), "Foam"; got != want {
		t.Fatalf("expected row header %s, got %s", want, got)
	}
}

func TestGotoInvalid(t *testing.T) {
	st := NewState()
	if err := st.Goto("ZZZ"); err == nil {
		t.Fatalf("expected error for missing row digits")
	}
}

func TestEditAndClipboard(t *testing.T) {
	st := NewState()
	st.EditCurrentCell("foo")
	if got := st.CurrentValue(); got != "foo" {
		t.Fatalf("expected foo, got %s", got)
	}
	st.YankCurrentCell()
	if st.Clipboard.Kind != ClipboardCell || st.Clipboard.CellValue != "foo" {
		t.Fatalf("clipboard mismatch: %+v", st.Clipboard)
	}
	st.CutCurrentCell()
	if val := st.CurrentValue(); val != "" {
		t.Fatalf("expected cell cleared after cut, got %q", val)
	}
	st.Clipboard.CellValue = "bar"
	if err := st.PasteClipboard(false); err != nil {
		t.Fatalf("paste failed: %v", err)
	}
	if got := st.CurrentValue(); got != "bar" {
		t.Fatalf("expected bar after paste, got %s", got)
	}
	st.Clipboard = Clipboard{}
	if err := st.PasteClipboard(false); err == nil {
		t.Fatalf("expected error when clipboard empty")
	}
}

func TestRowOperations(t *testing.T) {
	st := NewState()
	st.YankCurrentRow()
	if st.Clipboard.Kind != ClipboardRow || len(st.Clipboard.Rows) == 0 {
		t.Fatalf("expected row clipboard, got %+v", st.Clipboard)
	}
	st.CutCurrentRow()
	rows, _ := st.Workbook.MaxCoords()
	if rows != 4 {
		t.Fatalf("expected 4 rows after cut, got %d", rows)
	}
	if err := st.PasteClipboard(true); err != nil {
		t.Fatalf("row paste failed: %v", err)
	}
	rows, _ = st.Workbook.MaxCoords()
	if rows != 5 {
		t.Fatalf("expected 5 rows after paste, got %d", rows)
	}
	st.InsertRowAbove()
	rows, _ = st.Workbook.MaxCoords()
	if rows != 6 {
		t.Fatalf("expected 6 rows after insert, got %d", rows)
	}
	st.DeleteCurrentRow()
	rows, _ = st.Workbook.MaxCoords()
	if rows != 5 {
		t.Fatalf("expected row deletion to reduce count")
	}
}

func TestSelectionRangeClipboard(t *testing.T) {
	st := NewState()
	if err := st.Goto("A2"); err != nil {
		t.Fatalf("goto failed: %v", err)
	}
	st.BeginSelection(SelectionRange)
	st.Move(1, 1) // select A2:B3
	st.YankCurrentCell()
	if st.Clipboard.Kind != ClipboardRange {
		t.Fatalf("expected range clipboard, got %+v", st.Clipboard)
	}
	if len(st.Clipboard.RangeValues) != 2 || len(st.Clipboard.RangeValues[0]) != 2 {
		t.Fatalf("expected 2x2 range, got %+v", st.Clipboard.RangeValues)
	}
	if err := st.Goto("C4"); err != nil {
		t.Fatalf("goto failed: %v", err)
	}
	if err := st.PasteClipboard(false); err != nil {
		t.Fatalf("range paste failed: %v", err)
	}
	if got := st.Workbook.Cell(4, 3); got != "Foam" {
		t.Fatalf("expected pasted Foam at C4, got %s", got)
	}
	if got := st.Workbook.Cell(5, 4); got != "5" {
		t.Fatalf("expected pasted 5 at D5, got %s", got)
	}
}

func TestSelectionRowCutPaste(t *testing.T) {
	st := NewState()
	if err := st.Goto("A2"); err != nil {
		t.Fatalf("goto failed: %v", err)
	}
	st.BeginSelection(SelectionRow)
	st.Move(1, 0) // rows 2-3
	st.CutCurrentCell()
	if st.Clipboard.Kind != ClipboardRow || len(st.Clipboard.Rows) != 2 {
		t.Fatalf("expected two rows in clipboard, got %+v", st.Clipboard)
	}
	rows, _ := st.Workbook.MaxCoords()
	if rows != 3 {
		t.Fatalf("expected 3 rows after cutting two, got %d", rows)
	}
	if err := st.Goto("A2"); err != nil {
		t.Fatalf("goto failed: %v", err)
	}
	if err := st.PasteClipboard(false); err != nil {
		t.Fatalf("row paste failed: %v", err)
	}
	rows, _ = st.Workbook.MaxCoords()
	if rows != 5 {
		t.Fatalf("expected 5 rows after pasting back, got %d", rows)
	}
}

func TestSelectionSummary(t *testing.T) {
	st := NewState()
	st.BeginSelection(SelectionRange)
	st.Move(2, 1)
	if summary := st.SelectionSummary(); summary != "A1:B3 (3x2)" {
		t.Fatalf("unexpected summary for range: %s", summary)
	}
	st.ClearSelection()
	st.Goto("A2")
	st.BeginSelection(SelectionRow)
	st.Move(2, 0)
	if summary := st.SelectionSummary(); summary != "rows 2-4" {
		t.Fatalf("unexpected summary for rows: %s", summary)
	}
}

func TestPasteCellIntoSelection(t *testing.T) {
	st := NewState()
	st.Clipboard = Clipboard{Kind: ClipboardCell, CellValue: "X"}
	st.BeginSelection(SelectionRange)
	st.Move(1, 1)
	if err := st.PasteClipboard(false); err != nil {
		t.Fatalf("paste into selection failed: %v", err)
	}
	if got := st.Workbook.Cell(1, 1); got != "X" {
		t.Fatalf("expected A1 to be X, got %s", got)
	}
	if got := st.Workbook.Cell(2, 2); got != "X" {
		t.Fatalf("expected B2 to be X, got %s", got)
	}
}

func TestSearchForwardAndBackward(t *testing.T) {
	st := NewState()
	if err := st.Search("Foam", true); err != nil {
		t.Fatalf("forward search failed: %v", err)
	}
	if got := st.Address(); got != "A2" {
		t.Fatalf("expected search to land on A2, got %s", got)
	}
	// Backward search should wrap and find the header row via '?'
	if err := st.Search("Item", false); err != nil {
		t.Fatalf("reverse search failed: %v", err)
	}
	if got := st.Address(); got != "A1" {
		t.Fatalf("expected reverse search to land on A1, got %s", got)
	}
}

func TestSearchRepeat(t *testing.T) {
	st := NewState()
	if err := st.Search("$", true); err != nil {
		t.Fatalf("initial search failed: %v", err)
	}
	first := st.Address()
	if err := st.RepeatSearch(true); err != nil {
		t.Fatalf("repeat search (n) failed: %v", err)
	}
	second := st.Address()
	if first == second {
		t.Fatalf("repeat should move to next match")
	}
	if err := st.RepeatSearch(false); err != nil {
		t.Fatalf("reverse repeat (N) failed: %v", err)
	}
	if current := st.Address(); current != first {
		t.Fatalf("expected to return to previous match, got %s", current)
	}
	st.ClearSelection()
	st.lastSearchQuery = ""
	if err := st.RepeatSearch(true); err == nil {
		t.Fatalf("expected error when no previous search")
	}
}

func TestSearchCaseSensitivity(t *testing.T) {
	st := NewState()
	if err := st.Search("foam", true); err != nil {
		t.Fatalf("case-insensitive search should find result: %v", err)
	}
	st.SetSearchCaseSensitivity(true)
	if err := st.Search("foam", true); err == nil {
		t.Fatalf("case-sensitive search should fail for lowercase query")
	}
	if err := st.Search("Foam", true); err != nil {
		t.Fatalf("case-sensitive search should find exact case: %v", err)
	}
}

func TestStyleCopyPaste(t *testing.T) {
	wb := &workbook.Workbook{
		Cells: [][]string{
			{"A1", "B1"},
			{"A2", "B2"},
		},
		Sheet: "Sheet1",
		Styles: map[string]workbook.CellStyle{
			"A1": {Bold: true},
			"B1": {Italic: true},
		},
	}
	st := &State{
		Workbook: wb,
		Cursor:   Cursor{Row: 1, Col: 1},
	}
	if err := st.CopyStyle("A1:B1"); err != nil {
		t.Fatalf("copy style failed: %v", err)
	}
	if st.StyleClipboard.Width != 2 || st.StyleClipboard.Height != 1 {
		t.Fatalf("clipboard dimensions incorrect: %+v", st.StyleClipboard)
	}
	if err := st.PasteStyle("A2:B2"); err != nil {
		t.Fatalf("paste style failed: %v", err)
	}
	if style, ok := wb.Style("A2"); !ok || !style.Bold {
		t.Fatalf("expected bold style on A2")
	}
	if style, ok := wb.Style("B2"); !ok || !style.Italic {
		t.Fatalf("expected italic style on B2")
	}
	// Paste single-cell style across range
	if err := st.CopyStyle("A1"); err != nil {
		t.Fatalf("copy single style failed: %v", err)
	}
	if err := st.PasteStyle("A2:B2"); err != nil {
		t.Fatalf("paste single style to range failed: %v", err)
	}
	if style, ok := wb.Style("B2"); !ok || !style.Bold {
		t.Fatalf("expected bold style replicated on B2")
	}
}
