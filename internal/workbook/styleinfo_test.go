package workbook

import (
	"strings"
	"testing"
)

func TestCellStyleDescribe(t *testing.T) {
	cs := CellStyle{FillColor: "FF0000", FontColor: "00FF00", Bold: true, Italic: true}
	desc := cs.Describe()
	if desc == "" || desc == "no style information" {
		t.Fatalf("expected description, got %q", desc)
	}
	if !strings.Contains(desc, "red") || !strings.Contains(desc, "green") {
		t.Fatalf("expected color names, got %q", desc)
	}
}

func TestFuzzyColorName(t *testing.T) {
	name := fuzzyColorName("F08080")
	if name == "" {
		t.Fatalf("expected fuzzy color name")
	}
}
