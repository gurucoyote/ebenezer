package workbook

import (
	"fmt"
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
	name := colorName(hex)
	if name == "" {
		return fmt.Sprintf("#%s", hex)
	}
	return fmt.Sprintf("%s (#%s)", name, hex)
}

func colorName(hex string) string {
	switch hex {
	case "FF0000":
		return "red"
	case "00FF00":
		return "green"
	case "0000FF":
		return "blue"
	case "FFFF00":
		return "yellow"
	case "FFA500", "FF8C00":
		return "orange"
	case "800080", "9932CC":
		return "purple"
	case "FFC0CB", "FF69B4":
		return "pink"
	case "00FFFF":
		return "cyan"
	case "FFFFFF":
		return "white"
	case "000000":
		return "black"
	case "808080":
		return "gray"
	case "90EE90":
		return "light green"
	case "ADD8E6":
		return "light blue"
	case "FFB6C1":
		return "light pink"
	}
	return ""
}
