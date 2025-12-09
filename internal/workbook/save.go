package workbook

import (
	"fmt"
	"path/filepath"
	"sort"
	"strconv"
	"strings"

	excelize "github.com/xuri/excelize/v2"
)

func (w *Workbook) saveXLSX(path string) error {
	f := excelize.NewFile()
	sheetName := w.Sheet
	if sheetName == "" {
		sheetName = "Sheet1"
	}
	defaultSheet := f.GetSheetName(0)
	var targetSheet string
	if defaultSheet != sheetName {
		idx, err := f.NewSheet(sheetName)
		if err != nil {
			return fmt.Errorf("create sheet: %w", err)
		}
		if err := f.DeleteSheet(defaultSheet); err != nil {
			return fmt.Errorf("delete default sheet: %w", err)
		}
		f.SetActiveSheet(idx)
		targetSheet = sheetName
	} else {
		if err := f.SetSheetName(defaultSheet, sheetName); err != nil {
			return fmt.Errorf("rename sheet: %w", err)
		}
		targetSheet = sheetName
	}

	styleCache := map[string]int{}
	for rIdx, row := range w.Cells {
		for cIdx, val := range row {
			addr := fmt.Sprintf("%s%d", ColumnName(cIdx+1), rIdx+1)
			if err := f.SetCellStr(targetSheet, addr, val); err != nil {
				return fmt.Errorf("set cell %s: %w", addr, err)
			}
			if w.Styles != nil {
				if style, ok := w.Styles[strings.ToUpper(addr)]; ok && !style.Empty() {
					styleID, err := cachedStyleID(f, styleCache, style)
					if err != nil {
						return err
					}
					if err := f.SetCellStyle(targetSheet, addr, addr, styleID); err != nil {
						return fmt.Errorf("set style %s: %w", addr, err)
					}
				}
			}
		}
	}

	if cell := w.ActiveCell; cell != "" {
		_ = f.SetPanes(targetSheet, &excelize.Panes{
			Freeze:    false,
			Split:     false,
			Selection: []excelize.Selection{{ActiveCell: cell, SQRef: cell}},
		})
	}

	return f.SaveAs(path)
}

func cachedStyleID(f *excelize.File, cache map[string]int, cs CellStyle) (int, error) {
	key := styleCacheKey(cs)
	if id, ok := cache[key]; ok {
		return id, nil
	}
	es, err := buildExcelizeStyle(cs)
	if err != nil {
		return 0, err
	}
	id, err := f.NewStyle(es)
	if err != nil {
		return 0, fmt.Errorf("create style: %w", err)
	}
	cache[key] = id
	return id, nil
}

func styleCacheKey(cs CellStyle) string {
	var parts []string
	parts = append(parts, cs.FillColor, cs.FontColor, cs.NumberFormat, cs.HorizontalAlign, cs.VerticalAlign)
	parts = append(parts, strconv.FormatBool(cs.Bold), strconv.FormatBool(cs.Italic), strconv.FormatBool(cs.Underline))
	if len(cs.Borders) > 0 {
		keys := make([]string, 0, len(cs.Borders))
		for edge, style := range cs.Borders {
			keys = append(keys, fmt.Sprintf("%s:%s:%s", edge, style.Style, style.Color))
		}
		sort.Strings(keys)
		parts = append(parts, strings.Join(keys, ";"))
	}
	return strings.Join(parts, "|")
}

func buildExcelizeStyle(cs CellStyle) (*excelize.Style, error) {
	style := &excelize.Style{}
	if cs.FillColor != "" {
		style.Fill = excelize.Fill{Type: "pattern", Color: []string{"#" + cs.FillColor}, Pattern: 1}
	}
	if cs.FontColor != "" || cs.Bold || cs.Italic || cs.Underline {
		style.Font = &excelize.Font{Color: defaultColor(cs.FontColor), Bold: cs.Bold, Italic: cs.Italic}
		if cs.Underline {
			style.Font.Underline = "single"
		}
	}
	if cs.NumberFormat != "" {
		if strings.HasPrefix(cs.NumberFormat, "builtin:") {
			idStr := strings.TrimPrefix(cs.NumberFormat, "builtin:")
			if n, err := strconv.Atoi(idStr); err == nil {
				style.NumFmt = n
			}
		} else {
			fmtStr := cs.NumberFormat
			style.CustomNumFmt = &fmtStr
		}
	}
	if cs.HorizontalAlign != "" || cs.VerticalAlign != "" {
		style.Alignment = &excelize.Alignment{}
		if cs.HorizontalAlign != "" {
			style.Alignment.Horizontal = cs.HorizontalAlign
		}
		if cs.VerticalAlign != "" {
			style.Alignment.Vertical = cs.VerticalAlign
		}
	}
	if len(cs.Borders) > 0 {
		style.Border = []excelize.Border{}
		keys := make([]string, 0, len(cs.Borders))
		for edge := range cs.Borders {
			keys = append(keys, edge)
		}
		sort.Strings(keys)
		for _, edge := range keys {
			border := cs.Borders[edge]
			borderStyle := excelize.Border{Type: edge}
			borderStyle.Style = borderStyleFromName(border.Style)
			if strings.TrimSpace(border.Color) != "" {
				borderStyle.Color = defaultColor(border.Color)
			}
			style.Border = append(style.Border, borderStyle)
		}
	}
	return style, nil
}

func borderStyleFromName(name string) int {
	name = strings.ToLower(strings.TrimSpace(name))
	if strings.HasPrefix(name, "style-") {
		if n, err := strconv.Atoi(strings.TrimPrefix(name, "style-")); err == nil {
			return n
		}
	}
	if n, err := strconv.Atoi(name); err == nil {
		return n
	}
	for k, v := range borderStyleNames {
		if v == name {
			return k
		}
	}
	return 1 // default thin
}

func defaultColor(hex string) string {
	if strings.TrimSpace(hex) == "" {
		return "#000000"
	}
	if strings.HasPrefix(hex, "#") {
		return hex
	}
	return "#" + hex
}

// AddSheet creates a new sheet (blank or copied) inside an existing XLSX file.
func AddSheet(path, newName, copyFrom string) error {
	if strings.ToLower(filepath.Ext(path)) != ".xlsx" {
		return fmt.Errorf("sheet creation is only supported for .xlsx files")
	}
	if strings.TrimSpace(newName) == "" {
		return fmt.Errorf("sheet name is required")
	}
	f, err := excelize.OpenFile(path)
	if err != nil {
		return fmt.Errorf("open xlsx: %w", err)
	}
	defer f.Close()

	if sheetIndexCaseInsensitive(f, newName) >= 0 {
		return fmt.Errorf("sheet %s already exists", newName)
	}

	var targetIdx int
	trimCopy := strings.TrimSpace(copyFrom)
	if trimCopy != "" {
		sourceIdx := sheetIndexCaseInsensitive(f, trimCopy)
		if sourceIdx < 0 {
			return fmt.Errorf("source sheet %s not found", trimCopy)
		}
		targetIdx, err = f.NewSheet(newName)
		if err != nil {
			return fmt.Errorf("create sheet: %w", err)
		}
		if err := f.CopySheet(sourceIdx, targetIdx); err != nil {
			return fmt.Errorf("copy sheet: %w", err)
		}
	} else {
		targetIdx, err = f.NewSheet(newName)
		if err != nil {
			return fmt.Errorf("create sheet: %w", err)
		}
	}

	f.SetActiveSheet(targetIdx)
	if err := f.Save(); err != nil {
		return fmt.Errorf("save workbook: %w", err)
	}
	return nil
}

func sheetIndexCaseInsensitive(f *excelize.File, name string) int {
	if strings.TrimSpace(name) == "" {
		return -1
	}
	for _, sheet := range f.GetSheetList() {
		if strings.EqualFold(sheet, name) {
			idx, err := f.GetSheetIndex(sheet)
			if err != nil {
				return -1
			}
			return idx
		}
	}
	return -1
}
