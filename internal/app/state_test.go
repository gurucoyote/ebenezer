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
	if st.Clipboard.Kind != ClipboardRow || len(st.Clipboard.RowValues) == 0 {
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
