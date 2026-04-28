package workbook

import (
	"fmt"
	"sort"
	"strconv"
	"strings"
)

// SortKey defines a single sort column with direction.
type SortKey struct {
	Column    int  // 1-based column index
	Ascending bool // true for ascending, false for descending
}

// SortOpts configures a row sort operation.
type SortOpts struct {
	HeaderRows int       // number of header rows to keep fixed (default 1)
	StartRow   int       // 1-based start row; 0 = auto (HeaderRows+1)
	EndRow     int       // 1-based end row inclusive; 0 = auto (len(w.Cells))
	Keys       []SortKey // sort keys in priority order
}

// SortRows sorts data rows by one or more columns while keeping header rows
// fixed. Styles are remapped accordingly. Returns an error if validation fails.
func (w *Workbook) SortRows(opts SortOpts) error {
	if w == nil {
		return fmt.Errorf("no workbook data")
	}

	// Resolve defaults.
	if opts.HeaderRows < 0 {
		opts.HeaderRows = 0
	}
	if opts.HeaderRows == 0 {
		opts.HeaderRows = 1
	}
	if opts.StartRow == 0 {
		opts.StartRow = opts.HeaderRows + 1
	}
	if opts.EndRow == 0 {
		opts.EndRow = len(w.Cells)
	}

	// Validate.
	if opts.StartRow < 1 || opts.EndRow > len(w.Cells) || opts.StartRow > opts.EndRow {
		return fmt.Errorf("invalid row range %d-%d (workbook has %d rows)", opts.StartRow, opts.EndRow, len(w.Cells))
	}
	if len(opts.Keys) == 0 {
		return fmt.Errorf("at least one sort key is required")
	}
	for _, k := range opts.Keys {
		if k.Column < 1 {
			return fmt.Errorf("sort column must be >= 1")
		}
	}

	// Build slice of data rows with their original 1-based row index.
	type rowWithOrig struct {
		data []string
		orig int // 1-based original row
	}
	rows := make([]rowWithOrig, 0, opts.EndRow-opts.StartRow+1)
	for r := opts.StartRow; r <= opts.EndRow; r++ {
		rowCopy := make([]string, len(w.Cells[r-1]))
		copy(rowCopy, w.Cells[r-1])
		rows = append(rows, rowWithOrig{data: rowCopy, orig: r})
	}

	// Comparison helpers.
	less := func(a, b rowWithOrig) bool {
		for _, key := range opts.Keys {
			colIdx := key.Column - 1
			var av, bv string
			if colIdx < len(a.data) {
				av = a.data[colIdx]
			}
			if colIdx < len(b.data) {
				bv = b.data[colIdx]
			}
			// Try numeric comparison first.
			aNum, aErr := parseSortableNumber(av)
			bNum, bErr := parseSortableNumber(bv)
			if aErr == nil && bErr == nil {
				if aNum != bNum {
					if key.Ascending {
						return aNum < bNum
					}
					return aNum > bNum
				}
				continue
			}
			// Fallback to string comparison.
			cmp := strings.Compare(av, bv)
			if cmp != 0 {
				if key.Ascending {
					return cmp < 0
				}
				return cmp > 0
			}
		}
		return false // equal — preserve original order (stable sort)
	}

	// sort.Slice is not guaranteed stable, so we use sort.SliceStable.
	sort.SliceStable(rows, func(i, j int) bool {
		return less(rows[i], rows[j])
	})

	// Build row mapping: original 1-based row -> new 1-based row.
	rowMap := make(map[int]int, len(rows))
	for i, rw := range rows {
		rowMap[rw.orig] = opts.StartRow + i
	}

	// Rebuild Cells with sorted data rows.
	newCells := make([][]string, 0, len(w.Cells))
	newCells = append(newCells, w.Cells[:opts.StartRow-1]...)
	for _, rw := range rows {
		newCells = append(newCells, rw.data)
	}
	if opts.EndRow < len(w.Cells) {
		newCells = append(newCells, w.Cells[opts.EndRow:]...)
	}
	w.Cells = newCells

	// Remap styles for rows that moved.
	if w.Styles != nil {
		updated := make(map[string]CellStyle, len(w.Styles))
		for addr, style := range w.Styles {
			colLetters, row, err := splitAddress(addr)
			if err != nil {
				continue
			}
			if newRow, ok := rowMap[row]; ok {
				updated[fmt.Sprintf("%s%d", colLetters, newRow)] = style
				continue
			}
			updated[addr] = style
		}
		w.Styles = updated
	}

	return nil
}

// parseSortableNumber attempts to parse a string as a sortable number.
// It strips common currency symbols and handles comma/dot decimal separators.
func parseSortableNumber(s string) (float64, error) {
	s = strings.TrimSpace(s)
	if s == "" {
		return 0, fmt.Errorf("empty string")
	}
	// Strip common prefixes/suffixes.
	s = strings.TrimPrefix(s, "$")
	s = strings.TrimPrefix(s, "€")
	s = strings.TrimPrefix(s, "£")
	s = strings.TrimSuffix(s, "%")
	s = strings.TrimSpace(s)
	if s == "" {
		return 0, fmt.Errorf("only currency symbol")
	}
	// Normalize: if comma is used as decimal separator (e.g. "1,23"),
	// and there's no dot, replace comma with dot.
	if !strings.Contains(s, ".") && strings.Contains(s, ",") {
		// Check if comma is used as thousands separator (e.g. "1.234,56" or "1,234.56")
		// Heuristic: if there's a dot before the comma, it's European format.
		lastComma := strings.LastIndex(s, ",")
		if strings.Contains(s[:lastComma], ".") {
			// European thousands: "1.234,56" -> remove dots, replace comma with dot
			s = strings.ReplaceAll(s, ".", "")
			s = strings.Replace(s, ",", ".", 1)
		} else if lastComma == len(s)-3 || lastComma == len(s)-2 {
			// likely decimal separator: "1,23" or "1,2"
			s = strings.Replace(s, ",", ".", 1)
		} else {
			// likely thousands separator: "1,234"
			s = strings.ReplaceAll(s, ",", "")
		}
	} else if strings.Contains(s, ".") && strings.Contains(s, ",") {
		// "1,234.56" US format
		s = strings.ReplaceAll(s, ",", "")
	}
	return strconv.ParseFloat(s, 64)
}
