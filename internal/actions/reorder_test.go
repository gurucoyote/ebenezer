package actions

import (
	"strings"
	"testing"

	"ebenezer/internal/app"
)

func TestTableReorderActionMetadata(t *testing.T) {
	meta := TableReorder.Metadata()
	if meta.Name != "table-reorder" {
		t.Fatalf("expected name 'table-reorder', got %q", meta.Name)
	}
	if meta.Category != "editing" {
		t.Fatalf("expected category 'editing', got %q", meta.Category)
	}
	if meta.Idempotent {
		t.Fatal("table-reorder should not be idempotent")
	}
	if meta.Description == "" {
		t.Fatal("description must not be empty")
	}
}

func TestTableReorderActionBasic(t *testing.T) {
	// Build a state with a todo-like table.
	st := app.NewState()
	// Overwrite sample data with our test table.
	st.Workbook.Cells = [][]string{
		{"Datum", "Erledigt", "Todo", "Notiz"},
		{"2026-04-01", "", "Buy milk", ""},
		{"2026-04-02", "X", "Call Bob", "urgent"},
		{"2026-04-03", "", "Send report", ""},
		{"2026-04-04", "x", "Fix bug", "tracker#42"},
	}

	ctx := NewContext(st, nil)

	res, err := TableReorder.Exec(ctx, []string{"2", "X", "equals_ignore_case", "separator_rows=1"})
	if err != nil {
		t.Fatalf("Exec: %v", err)
	}
	if !strings.Contains(res.Message, "reordered") {
		t.Fatalf("expected reorder message, got %q", res.Message)
	}
	if !st.IsDirty() {
		t.Fatal("expected state to be dirty after reorder")
	}

	// Verify order: header, non-matching (Buy milk, Send report), separator, matching (Call Bob, Fix bug)
	if st.Workbook.Cells[1][2] != "Buy milk" {
		t.Errorf("row 1: expected 'Buy milk', got %q", st.Workbook.Cells[1][2])
	}
	if st.Workbook.Cells[2][2] != "Send report" {
		t.Errorf("row 2: expected 'Send report', got %q", st.Workbook.Cells[2][2])
	}
	if st.Workbook.Cells[4][2] != "Call Bob" {
		t.Errorf("row 4: expected 'Call Bob', got %q", st.Workbook.Cells[4][2])
	}
	if st.Workbook.Cells[5][2] != "Fix bug" {
		t.Errorf("row 5: expected 'Fix bug', got %q", st.Workbook.Cells[5][2])
	}
}

func TestTableReorderActionIsBlank(t *testing.T) {
	st := app.NewState()
	st.Workbook.Cells = [][]string{
		{"Item", "Note"},
		{"A", "has note"},
		{"B", ""},
		{"C", "another"},
	}

	ctx := NewContext(st, nil)
	res, err := TableReorder.Exec(ctx, []string{"2", "predicate=is_blank"})
	if err != nil {
		t.Fatalf("Exec: %v", err)
	}
	if !strings.Contains(res.Message, "reordered") {
		t.Fatalf("expected reorder message, got %q", res.Message)
	}

	// Non-blank first (A, C), blank (B) at bottom
	if st.Workbook.Cells[1][0] != "A" {
		t.Errorf("row 1: expected 'A', got %q", st.Workbook.Cells[1][0])
	}
	if st.Workbook.Cells[2][0] != "C" {
		t.Errorf("row 2: expected 'C', got %q", st.Workbook.Cells[2][0])
	}
	if st.Workbook.Cells[3][0] != "B" {
		t.Errorf("row 3: expected 'B', got %q", st.Workbook.Cells[3][0])
	}
}

func TestTableReorderActionMissingColumn(t *testing.T) {
	st := app.NewState()
	ctx := NewContext(st, nil)

	_, err := TableReorder.Exec(ctx, []string{})
	if err == nil || !strings.Contains(err.Error(), "column") {
		t.Fatalf("expected column error, got %v", err)
	}
}

func TestTableReorderActionMissingValue(t *testing.T) {
	st := app.NewState()
	ctx := NewContext(st, nil)

	_, err := TableReorder.Exec(ctx, []string{"2", "predicate=equals"})
	if err == nil || !strings.Contains(err.Error(), "value") {
		t.Fatalf("expected value error, got %v", err)
	}
}

func TestTableReorderActionColumnLetter(t *testing.T) {
	st := app.NewState()
	st.Workbook.Cells = [][]string{
		{"Item", "Status"},
		{"A", "done"},
		{"B", "todo"},
	}

	ctx := NewContext(st, nil)
	_, err := TableReorder.Exec(ctx, []string{"B", "done"})
	if err != nil {
		t.Fatalf("Exec with column letter: %v", err)
	}
}

func TestTableReorderDiscoverable(t *testing.T) {
	found := false
	for _, dto := range Discover() {
		if dto.Name == "table-reorder" {
			found = true
			break
		}
	}
	if !found {
		t.Fatal("table-reorder not found in Discover()")
	}
}

func TestParseReorderArgsBasic(t *testing.T) {
	result, err := parseReorderArgs([]string{"2", "X"})
	if err != nil {
		t.Fatalf("parseReorderArgs: %v", err)
	}
	if result.spOpts.KeyColumn != 2 {
		t.Errorf("KeyColumn: expected 2, got %d", result.spOpts.KeyColumn)
	}
	if result.spOpts.Value != "X" {
		t.Errorf("Value: expected 'X', got %q", result.spOpts.Value)
	}
	if result.spOpts.Predicate != "equals_ignore_case" {
		t.Errorf("Predicate: expected 'equals_ignore_case', got %q", result.spOpts.Predicate)
	}
}

func TestParseReorderArgsWithFlags(t *testing.T) {
	result, err := parseReorderArgs([]string{"2", "X", "separator_rows=2", "header_rows=2"})
	if err != nil {
		t.Fatalf("parseReorderArgs: %v", err)
	}
	if result.spOpts.SeparatorRows != 2 {
		t.Errorf("SeparatorRows: expected 2, got %d", result.spOpts.SeparatorRows)
	}
	if result.spOpts.HeaderRows != 2 {
		t.Errorf("HeaderRows: expected 2, got %d", result.spOpts.HeaderRows)
	}
}
