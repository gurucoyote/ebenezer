package workbook

import (
	"strings"
	"testing"
)

func TestStablePartitionMovesCompletedToBottom(t *testing.T) {
	// Setup: Datum | Erledigt | Todo | Notiz
	wb := &Workbook{
		Cells: [][]string{
			{"Datum", "Erledigt", "Todo", "Notiz"},
			{"2026-04-01", "", "Buy milk", ""},
			{"2026-04-02", "X", "Call Bob", "urgent"},
			{"2026-04-03", "", "Send report", ""},
			{"2026-04-04", "x", "Fix bug", "tracker#42"},
			{"2026-04-05", "", "Plan meeting", ""},
		},
		Sheet:  "Sheet1",
		Styles: map[string]CellStyle{},
	}

	_, err := wb.StablePartition(StablePartitionOpts{
		HeaderRows:       1,
		KeyColumn:        2,
		Predicate:        "equals_ignore_case",
		Value:            "X",
		MatchingPosition: "bottom",
		SeparatorRows:    1,
	})
	if err != nil {
		t.Fatalf("StablePartition: %v", err)
	}

	// Header must be untouched
	if !equalRows(wb.Cells[0], []string{"Datum", "Erledigt", "Todo", "Notiz"}) {
		t.Fatalf("header changed: %v", wb.Cells[0])
	}

	// Non-matching rows first (rows 2,4,6 in original = indices 1,3,5)
	// They should appear in original order: Buy milk, Send report, Plan meeting
	if wb.Cells[1][2] != "Buy milk" {
		t.Errorf("row 1 after header: expected 'Buy milk', got %q", wb.Cells[1][2])
	}
	if wb.Cells[2][2] != "Send report" {
		t.Errorf("row 2 after header: expected 'Send report', got %q", wb.Cells[2][2])
	}
	if wb.Cells[3][2] != "Plan meeting" {
		t.Errorf("row 3 after header: expected 'Plan meeting', got %q", wb.Cells[3][2])
	}

	// Separator row (blank)
	if !isBlankRow(wb.Cells[4]) {
		t.Errorf("row 4: expected blank separator, got %v", wb.Cells[4])
	}

	// Matching rows at bottom, in original order: Call Bob, Fix bug
	if wb.Cells[5][2] != "Call Bob" {
		t.Errorf("row 5: expected 'Call Bob', got %q", wb.Cells[5][2])
	}
	if wb.Cells[6][2] != "Fix bug" {
		t.Errorf("row 6: expected 'Fix bug', got %q", wb.Cells[6][2])
	}

	// Matching rows should carry their other columns
	if wb.Cells[5][3] != "urgent" {
		t.Errorf("Call Bob Notiz: expected 'urgent', got %q", wb.Cells[5][3])
	}
	if wb.Cells[6][3] != "tracker#42" {
		t.Errorf("Fix bug Notiz: expected 'tracker#42', got %q", wb.Cells[6][3])
	}
}

func TestStablePartitionMovesToTop(t *testing.T) {
	wb := &Workbook{
		Cells: [][]string{
			{"Item", "Status"},
			{"A", "done"},
			{"B", "todo"},
			{"C", "done"},
		},
		Sheet:  "Sheet1",
		Styles: map[string]CellStyle{},
	}

	_, err := wb.StablePartition(StablePartitionOpts{
		HeaderRows:       1,
		KeyColumn:        2,
		Predicate:        "equals",
		Value:            "done",
		MatchingPosition: "top",
		SeparatorRows:    0,
	})
	if err != nil {
		t.Fatalf("StablePartition: %v", err)
	}

	// Matching rows first: A, C
	if wb.Cells[1][1] != "done" {
		t.Errorf("row 1: expected 'done', got %q", wb.Cells[1][1])
	}
	if wb.Cells[1][0] != "A" {
		t.Errorf("row 1 item: expected 'A', got %q", wb.Cells[1][0])
	}
	if wb.Cells[2][0] != "C" {
		t.Errorf("row 2 item: expected 'C', got %q", wb.Cells[2][0])
	}
	// Non-matching: B
	if wb.Cells[3][0] != "B" {
		t.Errorf("row 3 item: expected 'B', got %q", wb.Cells[3][0])
	}
}

func TestStablePartitionPreservesStyles(t *testing.T) {
	wb := &Workbook{
		Cells: [][]string{
			{"Name", "Done"},
			{"Task1", "X"},
			{"Task2", ""},
			{"Task3", "X"},
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

	_, err := wb.StablePartition(StablePartitionOpts{
		HeaderRows:       1,
		KeyColumn:        2,
		Predicate:        "equals_ignore_case",
		Value:            "X",
		MatchingPosition: "bottom",
		SeparatorRows:    0,
	})
	if err != nil {
		t.Fatalf("StablePartition: %v", err)
	}

	// After partition: Task2 (green) is row 2, Task1 (red) is row 3, Task3 (blue) is row 4
	// Row 2 was originally row 3 (green)
	if s, ok := wb.Styles["A2"]; !ok || s.FillColor != "green" {
		t.Errorf("A2 style: expected green, got %v", s)
	}
	// Row 3 was originally row 2 (red)
	if s, ok := wb.Styles["A3"]; !ok || s.FillColor != "red" {
		t.Errorf("A3 style: expected red, got %v", s)
	}
	// Row 4 was originally row 4 (blue)
	if s, ok := wb.Styles["A4"]; !ok || s.FillColor != "blue" {
		t.Errorf("A4 style: expected blue, got %v", s)
	}
}

func TestStablePartitionIsBlank(t *testing.T) {
	wb := &Workbook{
		Cells: [][]string{
			{"Item", "Note"},
			{"A", "has note"},
			{"B", ""},
			{"C", "another"},
		},
		Sheet:  "Sheet1",
		Styles: map[string]CellStyle{},
	}

	_, err := wb.StablePartition(StablePartitionOpts{
		HeaderRows:       1,
		KeyColumn:        2,
		Predicate:        "is_blank",
		MatchingPosition: "bottom",
	})
	if err != nil {
		t.Fatalf("StablePartition: %v", err)
	}

	// Non-blank first: A, C. Blank: B
	if wb.Cells[1][0] != "A" {
		t.Errorf("row 1: expected 'A', got %q", wb.Cells[1][0])
	}
	if wb.Cells[2][0] != "C" {
		t.Errorf("row 2: expected 'C', got %q", wb.Cells[2][0])
	}
	if wb.Cells[3][0] != "B" {
		t.Errorf("row 3: expected 'B', got %q", wb.Cells[3][0])
	}
}

func TestStablePartitionInvalidKeyColumn(t *testing.T) {
	wb := SampleWorkbook()
	_, err := wb.StablePartition(StablePartitionOpts{
		KeyColumn: 0,
		Predicate: "equals",
		Value:     "X",
	})
	if err == nil || !strings.Contains(err.Error(), "KeyColumn") {
		t.Fatalf("expected KeyColumn error, got %v", err)
	}
}

func TestStablePartitionInvalidPredicate(t *testing.T) {
	wb := SampleWorkbook()
	_, err := wb.StablePartition(StablePartitionOpts{
		KeyColumn: 1,
		Predicate: "contains",
		Value:     "X",
	})
	if err == nil || !strings.Contains(err.Error(), "predicate") {
		t.Fatalf("expected predicate error, got %v", err)
	}
}

func TestStablePartitionNilWorkbook(t *testing.T) {
	var wb *Workbook
	_, err := wb.StablePartition(StablePartitionOpts{
		KeyColumn: 1,
		Predicate: "equals",
		Value:     "X",
	})
	if err == nil {
		t.Fatal("expected error for nil workbook")
	}
}

func TestStablePartitionNoMatchAllSameGroup(t *testing.T) {
	wb := &Workbook{
		Cells: [][]string{
			{"Item", "Done"},
			{"A", ""},
			{"B", ""},
			{"C", ""},
		},
		Sheet:  "Sheet1",
		Styles: map[string]CellStyle{},
	}

	_, err := wb.StablePartition(StablePartitionOpts{
		HeaderRows:       1,
		KeyColumn:        2,
		Predicate:        "equals_ignore_case",
		Value:            "X",
		MatchingPosition: "bottom",
	})
	if err != nil {
		t.Fatalf("StablePartition: %v", err)
	}

	// No matching rows, so order should be unchanged
	if wb.Cells[1][0] != "A" {
		t.Errorf("row 1: expected 'A', got %q", wb.Cells[1][0])
	}
	if wb.Cells[2][0] != "B" {
		t.Errorf("row 2: expected 'B', got %q", wb.Cells[2][0])
	}
	if wb.Cells[3][0] != "C" {
		t.Errorf("row 3: expected 'C', got %q", wb.Cells[3][0])
	}
}

// Helper functions

func equalRows(a, b []string) bool {
	if len(a) != len(b) {
		return false
	}
	for i := range a {
		if a[i] != b[i] {
			return false
		}
	}
	return true
}

func isBlankRow(row []string) bool {
	for _, cell := range row {
		if strings.TrimSpace(cell) != "" {
			return false
		}
	}
	return true
}
