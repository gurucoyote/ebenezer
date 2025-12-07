package workbook

import (
	"fmt"
	"math"
	"strconv"
	"strings"
)

// CellStyle captures the subset of formatting we currently expose.
type CellStyle struct {
	FillColor string
	FontColor string
	Bold      bool
	Italic    bool
	Underline bool
}

// Empty returns true when no styling metadata is present.
func (c CellStyle) Empty() bool {
	return c.FillColor == "" && c.FontColor == "" && !c.Bold && !c.Italic && !c.Underline
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
	if len(parts) == 0 {
		return "no style information"
	}
	return strings.Join(parts, ", ")
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
