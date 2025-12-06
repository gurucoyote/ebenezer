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
		if got := columnName(col); got != want {
			t.Fatalf("col %d: expected %s, got %s", col, want, got)
		}
	}
}
