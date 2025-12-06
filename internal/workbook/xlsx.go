package workbook

import (
	"fmt"

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

	return &Workbook{
		Cells: rows,
		Name:  path,
		Sheet: sheetName,
	}, nil
}
