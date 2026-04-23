package workbook

import (
	"encoding/json"
	"fmt"
	"math"
	"sort"
	"strconv"
	"strings"
)

// CellStyle captures the subset of formatting we currently expose.
type CellStyle struct {
	FillColor       string                 `json:"fillColor,omitempty"`
	FontColor       string                 `json:"fontColor,omitempty"`
	NumberFormat    string                 `json:"numberFormat,omitempty"`
	HorizontalAlign string                 `json:"horizontalAlign,omitempty"`
	VerticalAlign   string                 `json:"verticalAlign,omitempty"`
	Borders         map[string]BorderStyle `json:"borders,omitempty"`
	Bold            bool                   `json:"bold,omitempty"`
	Italic          bool                   `json:"italic,omitempty"`
	Underline       bool                   `json:"underline,omitempty"`
}

// BorderStyle summarizes a border edge for a cell.
type BorderStyle struct {
	Style string `json:"style,omitempty"`
	Color string `json:"color,omitempty"`
}

// Empty returns true when no styling metadata is present.
func (c CellStyle) Empty() bool {
	return c.FillColor == "" && c.FontColor == "" && c.NumberFormat == "" &&
		c.HorizontalAlign == "" && c.VerticalAlign == "" && len(c.Borders) == 0 &&
		!c.Bold && !c.Italic && !c.Underline
}

// cellStyleAlias is an alias type used inside UnmarshalJSON to avoid
// infinite recursion when calling json.Unmarshal on the standard struct.
type cellStyleAlias CellStyle

// UnmarshalJSON implements custom unmarshalling for CellStyle.
// It maps legacy field names (bg_color, font_color) to the canonical
// Go fields (FillColor, FontColor) so that both old and new JSON keys
// are accepted.
func (c *CellStyle) UnmarshalJSON(data []byte) error {
	// 1. Unmarshal using the standard json tags via the alias type.
	var alias cellStyleAlias
	if err := json.Unmarshal(data, &alias); err != nil {
		return err
	}

	// 2. Inspect a raw map for legacy field names.
	var raw map[string]json.RawMessage
	if err := json.Unmarshal(data, &raw); err != nil {
		// data was already valid JSON above, so this should not happen,
		// but fall back gracefully.
		*c = CellStyle(alias)
		return nil
	}

	if alias.FillColor == "" {
		if v, ok := raw["bg_color"]; ok {
			var s string
			if err := json.Unmarshal(v, &s); err == nil {
				alias.FillColor = s
			}
		}
	}
	if alias.FontColor == "" {
		if v, ok := raw["font_color"]; ok {
			var s string
			if err := json.Unmarshal(v, &s); err == nil {
				alias.FontColor = s
			}
		}
	}

	*c = CellStyle(alias)
	return nil
}

// Describe returns a human-friendly summary of the style.
func (c CellStyle) Describe() string {
	var parts []string
	if c.FillColor != "" {
		parts = append(parts, fmt.Sprintf("fill %s", describeColor(c.FillColor)))
	}
	if c.FontColor != "" {
		parts = append(parts, fmt.Sprintf("font %s", describeColor(c.FontColor)))
	}
	if c.Bold {
		parts = append(parts, "bold")
	}
	if c.Italic {
		parts = append(parts, "italic")
	}
	if c.Underline {
		parts = append(parts, "underline")
	}
	if c.NumberFormat != "" {
		parts = append(parts, fmt.Sprintf("numfmt %s", c.NumberFormat))
	}
	if len(c.Borders) > 0 {
		edges := make([]string, 0, len(c.Borders))
		for edge, border := range c.Borders {
			if border.Style != "" {
				edges = append(edges, fmt.Sprintf("%s(%s)", edge, border.Style))
			} else {
				edges = append(edges, edge)
			}
		}
		sort.Strings(edges)
		parts = append(parts, fmt.Sprintf("borders %s", strings.Join(edges, ",")))
	}
	if c.HorizontalAlign != "" || c.VerticalAlign != "" {
		parts = append(parts, fmt.Sprintf("align H=%s V=%s", emptySafe(c.HorizontalAlign), emptySafe(c.VerticalAlign)))
	}
	if len(parts) == 0 {
		return "no style information"
	}
	return strings.Join(parts, ", ")
}

func emptySafe(val string) string {
	if strings.TrimSpace(val) == "" {
		return "default"
	}
	return val
}

func describeColor(hex string) string {
	hex = strings.TrimSpace(strings.TrimPrefix(hex, "#"))
	hex = strings.ToUpper(hex)
	if len(hex) == 8 {
		hex = hex[2:]
	}
	name := fuzzyColorName(hex)
	if name == "" {
		return fmt.Sprintf("#%s", hex)
	}
	return fmt.Sprintf("%s (#%s)", name, hex)
}

type namedColor struct {
	name string
	hex  string
	lab  labColor
}

type labColor struct {
	L, A, B float64
}

var baseColors = []namedColor{
	{"red", "FF0000", labFromHex("FF0000")},
	{"orange", "FFA500", labFromHex("FFA500")},
	{"amber", "FFBF00", labFromHex("FFBF00")},
	{"yellow", "FFFF00", labFromHex("FFFF00")},
	{"lime", "BFFF00", labFromHex("BFFF00")},
	{"green", "00FF00", labFromHex("00FF00")},
	{"teal", "008080", labFromHex("008080")},
	{"cyan", "00FFFF", labFromHex("00FFFF")},
	{"sky blue", "87CEEB", labFromHex("87CEEB")},
	{"blue", "0000FF", labFromHex("0000FF")},
	{"indigo", "4B0082", labFromHex("4B0082")},
	{"purple", "800080", labFromHex("800080")},
	{"pink", "FF69B4", labFromHex("FF69B4")},
	{"light pink", "FFB6C1", labFromHex("FFB6C1")},
	{"light green", "90EE90", labFromHex("90EE90")},
	{"light blue", "ADD8E6", labFromHex("ADD8E6")},
	{"brown", "8B4513", labFromHex("8B4513")},
	{"gray", "808080", labFromHex("808080")},
	{"black", "000000", labFromHex("000000")},
	{"white", "FFFFFF", labFromHex("FFFFFF")},
}

func fuzzyColorName(hex string) string {
	if hex == "" {
		return ""
	}
	if c := exactColorName(hex); c != "" {
		return c
	}
	lab, ok := labFromHexSafe(hex)
	if !ok {
		return ""
	}
	best := math.MaxFloat64
	name := ""
	for _, candidate := range baseColors {
		delta := deltaE(lab, candidate.lab)
		if delta < best {
			best = delta
			name = candidate.name
		}
	}
	return name
}

func exactColorName(hex string) string {
	hex = strings.ToUpper(hex)
	for _, c := range baseColors {
		if c.hex == hex {
			return c.name
		}
	}
	return ""
}

func labFromHexSafe(hex string) (labColor, bool) {
	lab := labFromHex(hex)
	return lab, lab != (labColor{})
}

func labFromHex(hex string) labColor {
	hex = strings.TrimPrefix(hex, "#")
	if len(hex) != 6 {
		return labColor{}
	}
	r, err1 := strconv.ParseInt(hex[0:2], 16, 64)
	g, err2 := strconv.ParseInt(hex[2:4], 16, 64)
	b, err3 := strconv.ParseInt(hex[4:6], 16, 64)
	if err1 != nil || err2 != nil || err3 != nil {
		return labColor{}
	}
	return rgbToLab(float64(r)/255, float64(g)/255, float64(b)/255)
}

func rgbToLab(r, g, b float64) labColor {
	r = srgbToLinear(r)
	g = srgbToLinear(g)
	b = srgbToLinear(b)
	X := r*0.4124564 + g*0.3575761 + b*0.1804375
	Y := r*0.2126729 + g*0.7151522 + b*0.072175
	Z := r*0.0193339 + g*0.119192 + b*0.9503041
	return labColor{
		L: 116*labF(Y/0.95047) - 16,
		A: 500 * (labF(X/0.95047) - labF(Y/1.0)),
		B: 200 * (labF(Y/1.0) - labF(Z/1.08883)),
	}
}

func srgbToLinear(c float64) float64 {
	if c <= 0.04045 {
		return c / 12.92
	}
	return math.Pow((c+0.055)/1.055, 2.4)
}

func labF(t float64) float64 {
	if t > 0.008856 {
		return math.Cbrt(t)
	}
	return 7.787*t + 16.0/116.0
}

func deltaE(a, b labColor) float64 {
	dL := a.L - b.L
	dA := a.A - b.A
	dB := a.B - b.B
	return math.Sqrt(dL*dL + dA*dA + dB*dB)
}
