package workbook

import (
	"fmt"
	"strings"

	excelize "github.com/xuri/excelize/v2"
)

// FromXLSX loads cells from an .xlsx workbook. If sheet is empty, the first sheet
// is used.
func FromXLSX(path, sheet string) (*Workbook, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, fmt.Errorf("open xlsx: %w", err)
	}
	defer f.Close()

	sheetName := sheet
	if sheetName == "" {
		sheets := f.GetSheetList()
		if len(sheets) == 0 {
			return nil, fmt.Errorf("workbook has no sheets")
		}
		idx := f.GetActiveSheetIndex()
		if idx >= 0 && idx < len(sheets) {
			sheetName = sheets[idx]
		} else {
			sheetName = sheets[0]
		}
	}
	if sheetName == "" {
		return nil, fmt.Errorf("workbook has no sheets")
	}

	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("read sheet %s: %w", sheetName, err)
	}

	styles := map[string]CellStyle{}
	for rIdx, row := range rows {
		for cIdx := range row {
			addr := fmt.Sprintf("%s%d", ColumnName(cIdx+1), rIdx+1)
			style, err := extractCellStyle(f, sheetName, addr)
			if err == nil && !style.Empty() {
				styles[strings.ToUpper(addr)] = style
			}
		}
	}

	return &Workbook{
		Cells:  rows,
		Name:   path,
		Sheet:  sheetName,
		Styles: styles,
	}, nil
}

func extractCellStyle(f *excelize.File, sheet, axis string) (CellStyle, error) {
	idx, err := f.GetCellStyle(sheet, axis)
	if err != nil || idx == 0 {
		return CellStyle{}, err
	}
	style, err := f.GetStyle(idx)
	if err != nil {
		return CellStyle{}, err
	}
	var cs CellStyle
	if len(style.Fill.Color) > 0 {
		cs.FillColor = normalizeColor(style.Fill.Color[0])
	}
	if style.Font != nil {
		font := style.Font
		cs.FontColor = normalizeColor(font.Color)
		cs.Bold = font.Bold
		cs.Italic = font.Italic
		cs.Underline = font.Underline != ""
	}
	return cs, nil
}

func normalizeColor(color string) string {
	color = strings.TrimSpace(color)
	if strings.HasPrefix(color, "theme") {
		return ""
	}
	color = strings.TrimPrefix(color, "#")
	color = strings.ToUpper(color)
	if len(color) == 8 {
		color = color[2:]
	}
	return color
}
