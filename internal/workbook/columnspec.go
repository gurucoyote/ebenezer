package workbook

import (
	"fmt"
	"strconv"
	"strings"
)

// ParseColumnSpec converts inputs like "A", "B:D", "3", or "A1:D20" into
// inclusive 1-based column indexes.
func ParseColumnSpec(spec string) (int, int, error) {
	trimmed := strings.TrimSpace(strings.ToUpper(spec))
	if trimmed == "" {
		return 0, 0, fmt.Errorf("column spec required")
	}
	parts := strings.Split(trimmed, ":")
	if len(parts) > 2 {
		return 0, 0, fmt.Errorf("invalid column spec %s", spec)
	}
	start, err := columnIndexFromFragment(parts[0])
	if err != nil {
		return 0, 0, err
	}
	end := start
	if len(parts) == 2 {
		end, err = columnIndexFromFragment(parts[1])
		if err != nil {
			return 0, 0, err
		}
	}
	if start > end {
		start, end = end, start
	}
	return start, end, nil
}

func columnIndexFromFragment(fragment string) (int, error) {
	fragment = strings.TrimSpace(fragment)
	if fragment == "" {
		return 0, fmt.Errorf("invalid column fragment")
	}
	var letters strings.Builder
	var digits strings.Builder
	for _, r := range fragment {
		switch {
		case r >= 'A' && r <= 'Z':
			letters.WriteRune(r)
		case r >= '0' && r <= '9':
			digits.WriteRune(r)
		default:
			return 0, fmt.Errorf("invalid column fragment %s", fragment)
		}
	}
	if letters.Len() > 0 {
		return lettersToNumber(letters.String()), nil
	}
	if digits.Len() > 0 {
		n, err := strconv.Atoi(digits.String())
		if err != nil || n < 1 {
			return 0, fmt.Errorf("invalid column number %s", fragment)
		}
		return n, nil
	}
	return 0, fmt.Errorf("invalid column fragment %s", fragment)
}
