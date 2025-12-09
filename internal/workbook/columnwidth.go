package workbook

import (
	"math"
	"strings"
)

// DefaultColumnWidth mirrors Excel's default width when none is specified.
const DefaultColumnWidth = 8.43

// ColumnWidthOptions configures the auto-fit heuristic.
type ColumnWidthOptions struct {
	MinWidth        float64
	MaxWidth        float64
	Padding         float64
	MultilineBonus  float64
	CharacterFactor float64
}

// DefaultColumnWidthOptions defines the standard heuristic used by Ebenezer.
func DefaultColumnWidthOptions() ColumnWidthOptions {
	return ColumnWidthOptions{
		MinWidth:        10,
		MaxWidth:        70,
		Padding:         2.5,
		MultilineBonus:  2,
		CharacterFactor: 1.12,
	}
}

// ColumnWidth returns the stored width for the column (1-based).
func (w *Workbook) ColumnWidth(col int) (float64, bool) {
	if w == nil || col < 1 || w.ColumnWidths == nil {
		return 0, false
	}
	width, ok := w.ColumnWidths[col]
	return width, ok
}

// SetColumnWidth stores the given width for the column (1-based). Widths <= 0
// remove explicit overrides and fall back to the default width during save.
func (w *Workbook) SetColumnWidth(col int, width float64) {
	if w == nil || col < 1 {
		return
	}
	if w.ColumnWidths == nil {
		w.ColumnWidths = map[int]float64{}
	}
	if width <= 0 {
		delete(w.ColumnWidths, col)
		return
	}
	w.ColumnWidths[col] = width
}

// EstimateColumnWidth runs the heuristic across all rows for the column and
// returns the suggested width.
func (w *Workbook) EstimateColumnWidth(col int, opts ColumnWidthOptions) float64 {
	if w == nil {
		return opts.MinWidth
	}
	opts = normalizeWidthOptions(opts)
	longest := 0
	multiline := false
	for _, row := range w.Cells {
		if col-1 >= len(row) || col-1 < 0 {
			continue
		}
		cell := row[col-1]
		if cell == "" {
			continue
		}
		if strings.Contains(cell, "\n") {
			multiline = true
		}
		lines := strings.Split(cell, "\n")
		for _, line := range lines {
			trimmed := strings.TrimSpace(line)
			length := runeLen(trimmed)
			if length > longest {
				longest = length
			}
		}
	}
	width := float64(longest)*opts.CharacterFactor + opts.Padding
	if multiline {
		width += opts.MultilineBonus
	}
	if width < opts.MinWidth {
		width = opts.MinWidth
	}
	if width > opts.MaxWidth {
		width = opts.MaxWidth
	}
	return math.Round(width*10) / 10
}

func normalizeWidthOptions(opts ColumnWidthOptions) ColumnWidthOptions {
	if opts.MinWidth <= 0 {
		opts.MinWidth = DefaultColumnWidth
	}
	if opts.MaxWidth <= 0 || opts.MaxWidth < opts.MinWidth {
		opts.MaxWidth = opts.MinWidth
	}
	if opts.Padding < 0 {
		opts.Padding = 0
	}
	if opts.MultilineBonus < 0 {
		opts.MultilineBonus = 0
	}
	if opts.CharacterFactor <= 0 {
		opts.CharacterFactor = 1
	}
	return opts
}

func runeLen(s string) int {
	count := 0
	for range s {
		count++
	}
	return count
}
