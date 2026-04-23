package workbook

import (
	"encoding/json"
	"strings"
	"testing"
)

func TestCellStyleDescribe(t *testing.T) {
	cs := CellStyle{
		FillColor:       "FF0000",
		FontColor:       "00FF00",
		Bold:            true,
		Italic:          true,
		NumberFormat:    "#,##0",
		HorizontalAlign: "center",
		Borders: map[string]BorderStyle{
			"top": {Style: "thin", Color: "000000"},
		},
	}
	desc := cs.Describe()
	if desc == "" || desc == "no style information" {
		t.Fatalf("expected description, got %q", desc)
	}
	if !strings.Contains(desc, "red") || !strings.Contains(desc, "green") {
		t.Fatalf("expected color names, got %q", desc)
	}
	if !strings.Contains(desc, "numfmt") || !strings.Contains(desc, "borders") {
		t.Fatalf("expected number format and border info, got %q", desc)
	}
}

func TestFuzzyColorName(t *testing.T) {
	name := fuzzyColorName("F08080")
	if name == "" {
		t.Fatalf("expected fuzzy color name")
	}
}

func TestCellStyleUnmarshalJSON(t *testing.T) {
	tests := []struct {
		name      string
		json      string
		wantFill  string
		wantFont  string
	}{
		{
			name:     "legacy bg_color maps to FillColor",
			json:     `{"bg_color":"#FF0000"}`,
			wantFill: "#FF0000",
		},
		{
			name:     "legacy font_color maps to FontColor",
			json:     `{"font_color":"#00FF00"}`,
			wantFont: "#00FF00",
		},
		{
			name:     "canonical fillColor works",
			json:     `{"fillColor":"#FF0000"}`,
			wantFill: "#FF0000",
		},
		{
			name:     "canonical fontColor works",
			json:     `{"fontColor":"#00FF00"}`,
			wantFont: "#00FF00",
		},
		{
			name:     "canonical takes precedence over legacy",
			json:     `{"fillColor":"#111111","bg_color":"#222222","fontColor":"#333333","font_color":"#444444"}`,
			wantFill: "#111111",
			wantFont: "#333333",
		},
		{
			name:     "legacy fills in when canonical absent",
			json:     `{"bg_color":"#FF0000","bold":true}`,
			wantFill: "#FF0000",
		},
		{
			name: "both legacy fields together",
			json: `{"bg_color":"#FF0000","font_color":"#00FF00"}`,
			wantFill: "#FF0000",
			wantFont: "#00FF00",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			var cs CellStyle
			if err := json.Unmarshal([]byte(tt.json), &cs); err != nil {
				t.Fatalf("UnmarshalJSON failed: %v", err)
			}
			if cs.FillColor != tt.wantFill {
				t.Errorf("FillColor = %q, want %q", cs.FillColor, tt.wantFill)
			}
			if cs.FontColor != tt.wantFont {
				t.Errorf("FontColor = %q, want %q", cs.FontColor, tt.wantFont)
			}
		})
	}
}
