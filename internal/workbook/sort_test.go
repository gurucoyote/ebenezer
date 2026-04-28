package workbook

import (
	"testing"
)

func TestSortRowsBasic(t *testing.T) {
	wb := &Workbook{
		Cells: [][]string{
			{"Name", "Age"},
			{"Charlie", "30"},
			{"Alice", "25"},
			{"Bob", "35"},
		},
		Sheet:  "Sheet1",
		Styles: map[string]CellStyle{},
	}

	err := wb.SortRows(SortOpts{
		HeaderRows: 1,
		Keys:       []SortKey{{Column: 1, Ascending: true}},
	})
	if err != nil {
		t.Fatalf("SortRows: %v", err)
	}

	// Header unchanged
	if wb.Cells[0][0] != "Name" {
		t.Errorf("header changed: %v", wb.Cells[0])
	}
	// Sorted by Name ascending: Alice, Bob, Charlie
	if wb.Cells[1][0] != "Alice" {
		t.Errorf("row 1: expected Alice, got %q", wb.Cells[1][0])
	}
	if wb.Cells[2][0] != "Bob" {
		t.Errorf("row 2: expected Bob, got %q", wb.Cells[2][0])
	}
	if wb.Cells[3][0] != "Charlie" {
		t.Errorf("row 3: expected Charlie, got %q", wb.Cells[3][0])
	}
}

func TestSortRowsDesc(t *testing.T) {
	wb := &Workbook{
		Cells: [][]string{
			{"Name", "Age"},
			{"Charlie", "30"},
			{"Alice", "25"},
			{"Bob", "35"},
		},
		Sheet:  "Sheet1",
		Styles: map[string]CellStyle{},
	}

	err := wb.SortRows(SortOpts{
		HeaderRows: 1,
		Keys:       []SortKey{{Column: 2, Ascending: false}},
	})
	if err != nil {
		t.Fatalf("SortRows: %v", err)
	}

	// Sorted by Age descending: 35, 30, 25
	if wb.Cells[1][1] != "35" {
		t.Errorf("row 1 age: expected 35, got %q", wb.Cells[1][1])
	}
	if wb.Cells[2][1] != "30" {
		t.Errorf("row 2 age: expected 30, got %q", wb.Cells[2][1])
	}
	if wb.Cells[3][1] != "25" {
		t.Errorf("row 3 age: expected 25, got %q", wb.Cells[3][1])
	}
}

func TestSortRowsMultiKey(t *testing.T) {
	wb := &Workbook{
		Cells: [][]string{
			{"Dept", "Name", "Score"},
			{"Sales", "Alice", "90"},
			{"Sales", "Bob", "80"},
			{"IT", "Charlie", "95"},
			{"IT", "Dave", "85"},
			{"Sales", "Eve", "85"},
		},
		Sheet:  "Sheet1",
		Styles: map[string]CellStyle{},
	}

	err := wb.SortRows(SortOpts{
		HeaderRows: 1,
		Keys: []SortKey{
			{Column: 1, Ascending: true},  // Dept asc
			{Column: 3, Ascending: false}, // Score desc
		},
	})
	if err != nil {
		t.Fatalf("SortRows: %v", err)
	}

	// IT: Charlie(95), Dave(85)
	// Sales: Alice(90), Bob(80), Eve(85) — but wait, we sort by Score desc within Sales:
	// Alice(90), Eve(85), Bob(80)
	expected := []struct{ dept, name string }{
		{"IT", "Charlie"},
		{"IT", "Dave"},
		{"Sales", "Alice"},
		{"Sales", "Eve"},
		{"Sales", "Bob"},
	}
	for i, exp := range expected {
		if wb.Cells[i+1][0] != exp.dept || wb.Cells[i+1][1] != exp.name {
			t.Errorf("row %d: expected %s/%s, got %s/%s", i+1, exp.dept, exp.name, wb.Cells[i+1][0], wb.Cells[i+1][1])
		}
	}
}

func TestSortRowsPreservesStyles(t *testing.T) {
	wb := &Workbook{
		Cells: [][]string{
			{"Name", "Age"},
			{"Charlie", "30"},
			{"Alice", "25"},
			{"Bob", "35"},
		},
		Sheet: "Sheet1",
		Styles: map[string]CellStyle{
			"A2": {FillColor: "red"},
			"B2": {FillColor: "red"},
			"A3": {FillColor: "green"},
			"B3": {FillColor: "green"},
			"A4": {FillColor: "blue"},
			"B4": {FillColor: "blue"},
		},
	}

	err := wb.SortRows(SortOpts{
		HeaderRows: 1,
		Keys:       []SortKey{{Column: 1, Ascending: true}},
	})
	if err != nil {
		t.Fatalf("SortRows: %v", err)
	}

	// After sort: Alice(green) row2, Bob(blue) row3, Charlie(red) row4
	if s, ok := wb.Styles["A2"]; !ok || s.FillColor != "green" {
		t.Errorf("A2 style: expected green, got %v", s)
	}
	if s, ok := wb.Styles["A3"]; !ok || s.FillColor != "blue" {
		t.Errorf("A3 style: expected blue, got %v", s)
	}
	if s, ok := wb.Styles["A4"]; !ok || s.FillColor != "red" {
		t.Errorf("A4 style: expected red, got %v", s)
	}
}

func TestSortRowsNumeric(t *testing.T) {
	wb := &Workbook{
		Cells: [][]string{
			{"Item", "Price"},
			{"A", "$10"},
			{"B", "$2"},
			{"C", "$100"},
		},
		Sheet:  "Sheet1",
		Styles: map[string]CellStyle{},
	}

	err := wb.SortRows(SortOpts{
		HeaderRows: 1,
		Keys:       []SortKey{{Column: 2, Ascending: true}},
	})
	if err != nil {
		t.Fatalf("SortRows: %v", err)
	}

	// Numeric sort: 2, 10, 100
	if wb.Cells[1][0] != "B" || wb.Cells[1][1] != "$2" {
		t.Errorf("row 1: expected B/$2, got %s/%s", wb.Cells[1][0], wb.Cells[1][1])
	}
	if wb.Cells[2][0] != "A" || wb.Cells[2][1] != "$10" {
		t.Errorf("row 2: expected A/$10, got %s/%s", wb.Cells[2][0], wb.Cells[2][1])
	}
	if wb.Cells[3][0] != "C" || wb.Cells[3][1] != "$100" {
		t.Errorf("row 3: expected C/$100, got %s/%s", wb.Cells[3][0], wb.Cells[3][1])
	}
}

func TestSortRowsWithRange(t *testing.T) {
	wb := &Workbook{
		Cells: [][]string{
			{"Name", "Age"},
			{"Charlie", "30"},
			{"Alice", "25"},
			{"Bob", "35"},
			{"Diana", "28"},
		},
		Sheet:  "Sheet1",
		Styles: map[string]CellStyle{},
	}

	err := wb.SortRows(SortOpts{
		HeaderRows: 1,
		StartRow:   2,
		EndRow:     4,
		Keys:       []SortKey{{Column: 1, Ascending: true}},
	})
	if err != nil {
		t.Fatalf("SortRows: %v", err)
	}

	// Only rows 2-4 sorted: Alice, Bob, Charlie; Diana stays at row 5
	if wb.Cells[1][0] != "Alice" {
		t.Errorf("row 1: expected Alice, got %q", wb.Cells[1][0])
	}
	if wb.Cells[2][0] != "Bob" {
		t.Errorf("row 2: expected Bob, got %q", wb.Cells[2][0])
	}
	if wb.Cells[3][0] != "Charlie" {
		t.Errorf("row 3: expected Charlie, got %q", wb.Cells[3][0])
	}
	if wb.Cells[4][0] != "Diana" {
		t.Errorf("row 4: expected Diana, got %q", wb.Cells[4][0])
	}
}

func TestSortRowsNoKeysError(t *testing.T) {
	wb := SampleWorkbook()
	err := wb.SortRows(SortOpts{HeaderRows: 1})
	if err == nil {
		t.Fatal("expected error for missing keys")
	}
}

func TestSortRowsInvalidColumn(t *testing.T) {
	wb := SampleWorkbook()
	err := wb.SortRows(SortOpts{
		HeaderRows: 1,
		Keys:       []SortKey{{Column: 0, Ascending: true}},
	})
	if err == nil {
		t.Fatal("expected error for invalid column")
	}
}
