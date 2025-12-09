package app

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"ebenezer/internal/workbook"
)

// SaveRangeToFile persists the provided range data to a CSV or XLSX file.
func SaveRangeToFile(values [][]string, styles map[int]map[int]workbook.CellStyle, path string) (string, int64, error) {
	trimmed := strings.TrimSpace(path)
	if trimmed == "" {
		return "", 0, errors.New("path is required")
	}
	if len(values) == 0 {
		return "", 0, errors.New("no range data to save")
	}
	width := 0
	for _, row := range values {
		if len(row) > width {
			width = len(row)
		}
	}
	if width == 0 {
		return "", 0, errors.New("no range data to save")
	}
	ext := strings.ToLower(filepath.Ext(trimmed))
	if ext != ".csv" && ext != ".xlsx" {
		return "", 0, fmt.Errorf("unsupported extension %s", ext)
	}
	if dir := filepath.Dir(trimmed); dir != "." && dir != "" {
		if err := os.MkdirAll(dir, 0o755); err != nil {
			return "", 0, fmt.Errorf("create directory: %w", err)
		}
	}
	wb := &workbook.Workbook{
		Cells:      cloneRangeValues(values),
		Sheet:      "Selection",
		Name:       filepath.Base(trimmed),
		Styles:     convertRangeStyles(styles),
		ActiveCell: "A1",
	}
	if err := wb.Save(trimmed); err != nil {
		return "", 0, err
	}
	info, err := os.Stat(trimmed)
	if err != nil {
		return "", 0, err
	}
	return strings.TrimPrefix(ext, "."), info.Size(), nil
}

func cloneRangeValues(values [][]string) [][]string {
	dup := make([][]string, len(values))
	for i, row := range values {
		if row == nil {
			continue
		}
		copyRow := make([]string, len(row))
		copy(copyRow, row)
		dup[i] = copyRow
	}
	return dup
}

func convertRangeStyles(styles map[int]map[int]workbook.CellStyle) map[string]workbook.CellStyle {
	if len(styles) == 0 {
		return nil
	}
	converted := make(map[string]workbook.CellStyle)
	for r, cols := range styles {
		for c, style := range cols {
			if style.Empty() {
				continue
			}
			addr := fmt.Sprintf("%s%d", workbook.ColumnName(c+1), r+1)
			converted[strings.ToUpper(addr)] = style
		}
	}
	if len(converted) == 0 {
		return nil
	}
	return converted
}
