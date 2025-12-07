package workbook

import (
	"fmt"
	"strings"

	excelize "github.com/xuri/excelize/v2"
)

// FromXLSX loads cells from an .xlsx workbook. If sheet is empty, the first sheet
// is used. It also returns the sheet list and last active cell metadata.
func FromXLSX(path, sheet string) (*Workbook, []string, string, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, nil, "", fmt.Errorf("open xlsx: %w", err)
	}
	defer f.Close()

	sheets := f.GetSheetList()
	if len(sheets) == 0 {
		return nil, nil, "", fmt.Errorf("workbook has no sheets")
	}

	sheetName := sheet
	if sheetName == "" {
		idx := f.GetActiveSheetIndex()
		if idx >= 0 && idx < len(sheets) {
			sheetName = sheets[idx]
		} else {
			sheetName = sheets[0]
		}
	}
	if sheetName == "" {
		return nil, nil, "", fmt.Errorf("workbook has no sheets")
	}

	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, nil, "", fmt.Errorf("read sheet %s: %w", sheetName, err)
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

	activeCell := activeCellFromSheet(f, sheetName)

	return &Workbook{
		Cells:  rows,
		Name:   path,
		Sheet:  sheetName,
		Styles: styles,
	}, sheets, activeCell, nil
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

func activeCellFromSheet(f *excelize.File, sheet string) string {
	panes, err := f.GetPanes(sheet)
	if err == nil {
		for _, sel := range panes.Selection {
			if sel.ActiveCell != "" {
				return sel.ActiveCell
			}
		}
		if panes.TopLeftCell != "" {
			return panes.TopLeftCell
		}
	}
	return ""
}
