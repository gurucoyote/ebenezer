package workbook

import (
	"fmt"
	"strings"

	excelize "github.com/xuri/excelize/v2"
)

// RichTextWarning reports a cell that contains rich-text runs.
type RichTextWarning struct {
	Sheet string
	Cell  string
	Runs  int
}

// ScanRichText scans an XLSX file for cells that contain rich-text runs.
func ScanRichText(path, sheet string) ([]RichTextWarning, error) {
	if path == "" {
		return nil, fmt.Errorf("path is required")
	}
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, fmt.Errorf("open xlsx: %w", err)
	}
	sanitizeTheme(f)
	defer f.Close()

	sheets := f.GetSheetList()
	if len(sheets) == 0 {
		return nil, fmt.Errorf("workbook has no sheets")
	}

	targetSheets := sheets
	if sheet != "" {
		name, ok := findSheetName(sheets, sheet)
		if !ok {
			return nil, fmt.Errorf("sheet %s not found", sheet)
		}
		targetSheets = []string{name}
	}

	var warnings []RichTextWarning
	for _, sheetName := range targetSheets {
		rows, err := f.GetRows(sheetName)
		if err != nil {
			return nil, fmt.Errorf("read sheet %s: %w", sheetName, err)
		}
		for rIdx, row := range rows {
			for cIdx, value := range row {
				if value == "" {
					continue
				}
				cell := fmt.Sprintf("%s%d", ColumnName(cIdx+1), rIdx+1)
				runs, err := f.GetCellRichText(sheetName, cell)
				if err != nil {
					return nil, fmt.Errorf("read rich text %s!%s: %w", sheetName, cell, err)
				}
				if isRichTextRuns(runs) {
					warnings = append(warnings, RichTextWarning{
						Sheet: sheetName,
						Cell:  cell,
						Runs:  len(runs),
					})
				}
			}
		}
	}
	return warnings, nil
}

// RichTextRunMap converts warnings to an address->run-count map for a sheet.
func RichTextRunMap(warnings []RichTextWarning, sheet string) map[string]int {
	result := map[string]int{}
	if sheet == "" {
		return result
	}
	for _, warning := range warnings {
		if !strings.EqualFold(warning.Sheet, sheet) {
			continue
		}
		addr := strings.ToUpper(warning.Cell)
		result[addr] = warning.Runs
	}
	return result
}

func isRichTextRuns(runs []excelize.RichTextRun) bool {
	if len(runs) == 0 {
		return false
	}
	if len(runs) > 1 {
		return true
	}
	return runs[0].Font != nil
}

func findSheetName(sheets []string, target string) (string, bool) {
	for _, name := range sheets {
		if strings.EqualFold(name, target) {
			return name, true
		}
	}
	return "", false
}
