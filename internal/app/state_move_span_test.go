package app

import (
	"testing"

	"ebenezer/internal/workbook"
)

func TestMoveSpan_PositionalEmptyToData(t *testing.T) {
	wb := &workbook.Workbook{Cells: [][]string{{""}, {""}, {"A"}}}
	s := &State{Workbook: wb, Cursor: Cursor{Row: 1, Col: 1}}

	if err := s.MoveSpan("down"); err != nil {
		t.Fatalf("MoveSpan down: %v", err)
	}
	if s.Cursor.Row != 3 || s.Cursor.Col != 1 {
		t.Fatalf("expected jump to row3 col1, got %d,%d", s.Cursor.Row, s.Cursor.Col)
	}
}

func TestMoveSpan_WithinRunStopsAtEdge(t *testing.T) {
	wb := &workbook.Workbook{Cells: [][]string{{"A", "B", ""}}}
	s := &State{Workbook: wb, Cursor: Cursor{Row: 1, Col: 1}}

	if err := s.MoveSpan("right"); err != nil {
		t.Fatalf("MoveSpan right: %v", err)
	}
	if s.Cursor.Col != 2 {
		t.Fatalf("expected land on last filled col=2, got %d", s.Cursor.Col)
	}
}

func TestMoveSpan_FromEmptyAcrossEmpty(t *testing.T) {
	wb := &workbook.Workbook{Cells: [][]string{{"", "", "A"}}}
	s := &State{Workbook: wb, Cursor: Cursor{Row: 1, Col: 1}}

	if err := s.MoveSpan("right"); err != nil {
		t.Fatalf("MoveSpan right: %v", err)
	}
	if s.Cursor.Col != 3 {
		t.Fatalf("expected first non-empty col=3, got %d", s.Cursor.Col)
	}
}

func TestMoveSpan_InvalidDirection(t *testing.T) {
	wb := &workbook.Workbook{Cells: [][]string{{"A"}}}
	s := &State{Workbook: wb, Cursor: Cursor{Row: 1, Col: 1}}
	if err := s.MoveSpan("diag"); err == nil {
		t.Fatalf("expected error for invalid direction")
	}
}
