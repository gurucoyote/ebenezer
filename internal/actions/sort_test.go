package actions

import (
	"testing"

	"ebenezer/internal/app"
	"ebenezer/internal/workbook"
)

func TestTableSortActionBasic(t *testing.T) {
	state := app.NewState()
	state.Workbook = &workbook.Workbook{
		Cells: [][]string{
			{"Name", "Age"},
			{"Charlie", "30"},
			{"Alice", "25"},
			{"Bob", "35"},
		},
		Sheet:  "Sheet1",
		Styles: map[string]workbook.CellStyle{},
	}
	ctx := NewContext(state, nil)

	res, err := TableSort.Exec(ctx, []string{"A"})
	if err != nil {
		t.Fatalf("Exec: %v", err)
	}
	if res.Message == "" {
		t.Error("expected non-empty message")
	}

	// Check sorted order
	if state.Workbook.Cells[1][0] != "Alice" {
		t.Errorf("row 1: expected Alice, got %q", state.Workbook.Cells[1][0])
	}
	if state.Workbook.Cells[2][0] != "Bob" {
		t.Errorf("row 2: expected Bob, got %q", state.Workbook.Cells[2][0])
	}
	if state.Workbook.Cells[3][0] != "Charlie" {
		t.Errorf("row 3: expected Charlie, got %q", state.Workbook.Cells[3][0])
	}
}

func TestTableSortActionDesc(t *testing.T) {
	state := app.NewState()
	state.Workbook = &workbook.Workbook{
		Cells: [][]string{
			{"Name", "Age"},
			{"Charlie", "30"},
			{"Alice", "25"},
			{"Bob", "35"},
		},
		Sheet:  "Sheet1",
		Styles: map[string]workbook.CellStyle{},
	}
	ctx := NewContext(state, nil)

	_, err := TableSort.Exec(ctx, []string{"2:desc"})
	if err != nil {
		t.Fatalf("Exec: %v", err)
	}

	if state.Workbook.Cells[1][1] != "35" {
		t.Errorf("row 1 age: expected 35, got %q", state.Workbook.Cells[1][1])
	}
	if state.Workbook.Cells[2][1] != "30" {
		t.Errorf("row 2 age: expected 30, got %q", state.Workbook.Cells[2][1])
	}
	if state.Workbook.Cells[3][1] != "25" {
		t.Errorf("row 3 age: expected 25, got %q", state.Workbook.Cells[3][1])
	}
}

func TestTableSortActionMultiKey(t *testing.T) {
	state := app.NewState()
	state.Workbook = &workbook.Workbook{
		Cells: [][]string{
			{"Dept", "Name", "Score"},
			{"Sales", "Alice", "90"},
			{"Sales", "Bob", "80"},
			{"IT", "Charlie", "95"},
			{"IT", "Dave", "85"},
		},
		Sheet:  "Sheet1",
		Styles: map[string]workbook.CellStyle{},
	}
	ctx := NewContext(state, nil)

	_, err := TableSort.Exec(ctx, []string{"A:asc", "C:desc"})
	if err != nil {
		t.Fatalf("Exec: %v", err)
	}

	if state.Workbook.Cells[1][0] != "IT" || state.Workbook.Cells[1][1] != "Charlie" {
		t.Errorf("row 1: expected IT/Charlie, got %s/%s", state.Workbook.Cells[1][0], state.Workbook.Cells[1][1])
	}
	if state.Workbook.Cells[2][0] != "IT" || state.Workbook.Cells[2][1] != "Dave" {
		t.Errorf("row 2: expected IT/Dave, got %s/%s", state.Workbook.Cells[2][0], state.Workbook.Cells[2][1])
	}
}

func TestTableSortActionByFlag(t *testing.T) {
	state := app.NewState()
	state.Workbook = &workbook.Workbook{
		Cells: [][]string{
			{"Name", "Age"},
			{"Charlie", "30"},
			{"Alice", "25"},
			{"Bob", "35"},
		},
		Sheet:  "Sheet1",
		Styles: map[string]workbook.CellStyle{},
	}
	ctx := NewContext(state, nil)

	_, err := TableSort.Exec(ctx, []string{"by=B:desc"})
	if err != nil {
		t.Fatalf("Exec: %v", err)
	}

	if state.Workbook.Cells[1][1] != "35" {
		t.Errorf("row 1 age: expected 35, got %q", state.Workbook.Cells[1][1])
	}
}

func TestTableSortActionNoWorkbook(t *testing.T) {
	state := app.NewState()
	state.Workbook = nil
	ctx := NewContext(state, nil)

	_, err := TableSort.Exec(ctx, []string{"A"})
	if err == nil {
		t.Fatal("expected error when no workbook loaded")
	}
}

func TestParseSortKeys(t *testing.T) {
	cases := []struct {
		input    string
		expected []workbook.SortKey
	}{
		{"A", []workbook.SortKey{{Column: 1, Ascending: true}}},
		{"B:desc", []workbook.SortKey{{Column: 2, Ascending: false}}},
		{"2:asc", []workbook.SortKey{{Column: 2, Ascending: true}}},
		{"C:d", []workbook.SortKey{{Column: 3, Ascending: false}}},
		{"A:asc B:desc", []workbook.SortKey{{Column: 1, Ascending: true}, {Column: 2, Ascending: false}}},
	}

	for _, c := range cases {
		keys, err := parseSortKeys(c.input)
		if err != nil {
			t.Errorf("parseSortKeys(%q): unexpected error: %v", c.input, err)
			continue
		}
		if len(keys) != len(c.expected) {
			t.Errorf("parseSortKeys(%q): expected %d keys, got %d", c.input, len(c.expected), len(keys))
			continue
		}
		for i, k := range keys {
			if k.Column != c.expected[i].Column || k.Ascending != c.expected[i].Ascending {
				t.Errorf("parseSortKeys(%q)[%d]: expected %+v, got %+v", c.input, i, c.expected[i], k)
			}
		}
	}
}

func TestParseSortKeysInvalid(t *testing.T) {
	cases := []string{
		"0:asc",
		"A:up",
		":desc",
	}
	for _, input := range cases {
		_, err := parseSortKeys(input)
		if err == nil {
			t.Errorf("parseSortKeys(%q): expected error", input)
		}
	}
}
