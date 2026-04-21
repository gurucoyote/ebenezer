package workbook

import (
	"fmt"
	"strconv"
	"strings"
)

// ParseCellAddress parses a cell address like "A1" into 1-based row and column.
func ParseCellAddress(address string) (row, col int, err error) {
	addr := strings.TrimSpace(address)
	if addr == "" {
		return 0, 0, fmt.Errorf("address required")
	}
	addr = strings.ToUpper(addr)
	var letters, digits strings.Builder
	for _, r := range addr {
		switch {
		case r >= 'A' && r <= 'Z':
			letters.WriteRune(r)
		case r >= '0' && r <= '9':
			digits.WriteRune(r)
		default:
			return 0, 0, fmt.Errorf("invalid character %q in address", r)
		}
	}
	if letters.Len() == 0 || digits.Len() == 0 {
		return 0, 0, fmt.Errorf("address must include column letters and row digits")
	}
	col = lettersToNumber(letters.String())
	row, err = strconv.Atoi(digits.String())
	if err != nil {
		return 0, 0, fmt.Errorf("invalid row: %w", err)
	}
	return row, col, nil
}

// ParseRange parses an A1-style range string (e.g. "A1:D20") and returns
// 1-based row and column bounds, normalised so that start <= end.
func ParseRange(rangeStr string) (startRow, startCol, endRow, endCol int, err error) {
	parts := strings.Split(strings.TrimSpace(rangeStr), ":")
	if len(parts) > 2 {
		return 0, 0, 0, 0, fmt.Errorf("invalid range %q", rangeStr)
	}
	startRow, startCol, err = ParseCellAddress(strings.TrimSpace(parts[0]))
	if err != nil {
		return 0, 0, 0, 0, err
	}
	endRow, endCol = startRow, startCol
	if len(parts) == 2 {
		endRow, endCol, err = ParseCellAddress(strings.TrimSpace(parts[1]))
		if err != nil {
			return 0, 0, 0, 0, err
		}
	}
	if startRow > endRow {
		startRow, endRow = endRow, startRow
	}
	if startCol > endCol {
		startCol, endCol = endCol, startCol
	}
	return startRow, startCol, endRow, endCol, nil
}
