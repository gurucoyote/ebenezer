package actions

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"time"

	"ebenezer/internal/workbook"
)

var Info Action = infoAction{}

func init() {
	Register(Info)
}

type infoAction struct{}

type WorkbookInfo struct {
	Path        string
	ActiveSheet string
	Rows        int
	Cols        int
	SheetNames  []string
	ActiveCell  string
	FileSize    int64
	ModTime     time.Time
	Styles      StyleSummary
}

// StyleSummary aggregates formatting usage for the active workbook.
type StyleSummary struct {
	StyledCells    int            `json:"styledCells"`
	FillColors     map[string]int `json:"fillColors,omitempty"`
	FontColors     map[string]int `json:"fontColors,omitempty"`
	NumberFormats  map[string]int `json:"numberFormats,omitempty"`
	BorderUsage    map[string]int `json:"borderUsage,omitempty"`
	BoldCells      int            `json:"boldCells,omitempty"`
	ItalicCells    int            `json:"italicCells,omitempty"`
	UnderlineCells int            `json:"underlineCells,omitempty"`
}

func (infoAction) Name() string { return "info" }

func (infoAction) Metadata() Metadata {
	return Metadata{
		Name:        "info",
		Description: "Summarize workbook metadata (sheets, dimensions, cursor, file info)",
		Category:    "metadata",
		Args: []Arg{
			{Name: "path", Description: "Optional file to inspect", Optional: true},
			{Name: "--details", Description: "Include file size/timestamps", Optional: true},
		},
		Idempotent: true,
	}
}

func (infoAction) Exec(ctx Context, args []string) (Result, error) {
	if ctx.State == nil {
		return Result{}, errors.New("no workbook context")
	}
	var path string
	showDetails := false
	for _, arg := range args {
		if arg == "--details" {
			showDetails = true
			continue
		}
		if strings.HasPrefix(arg, "--details=") {
			value := strings.TrimPrefix(arg, "--details=")
			showDetails = value == "1" || strings.EqualFold(value, "true") || strings.EqualFold(value, "yes")
			continue
		}
		if path == "" {
			path = arg
		}
	}

	summary, err := buildWorkbookInfo(ctx, path)
	if err != nil {
		return Result{}, err
	}

	var b strings.Builder
	b.WriteString(fmt.Sprintf("path: %s\n", displayPath(summary.Path)))
	b.WriteString(fmt.Sprintf("sheet: %s (rows=%d cols=%d)\n", summary.ActiveSheet, summary.Rows, summary.Cols))
	if summary.ActiveCell != "" {
		b.WriteString(fmt.Sprintf("active cell: %s\n", summary.ActiveCell))
	}
	if len(summary.SheetNames) > 0 {
		b.WriteString("sheets:\n")
		for _, name := range summary.SheetNames {
			marker := " "
			if strings.EqualFold(name, summary.ActiveSheet) {
				marker = "*"
			}
			b.WriteString(fmt.Sprintf(" %s %s\n", marker, name))
		}
	}
	if showDetails {
		b.WriteString(fmt.Sprintf("file size: %d bytes\n", summary.FileSize))
		if !summary.ModTime.IsZero() {
			b.WriteString(fmt.Sprintf("modified: %s\n", summary.ModTime.Format(time.RFC3339)))
		}
		if summary.Styles.StyledCells > 0 {
			b.WriteString(fmt.Sprintf("styled cells: %d\n", summary.Styles.StyledCells))
			if len(summary.Styles.FillColors) > 0 {
				b.WriteString(fmt.Sprintf("fill palettes: %d unique\n", len(summary.Styles.FillColors)))
			}
			if len(summary.Styles.NumberFormats) > 0 {
				b.WriteString(fmt.Sprintf("number formats: %d unique\n", len(summary.Styles.NumberFormats)))
			}
		}
	}

	return Result{Message: b.String(), Data: summary}, nil
}

func buildWorkbookInfo(ctx Context, path string) (WorkbookInfo, error) {
	if strings.TrimSpace(path) != "" {
		wb, sheets, active, err := workbook.FromFile(path, "")
		if err != nil {
			return WorkbookInfo{}, err
		}
		rows, cols := wb.MaxCoords()
		size, mod := fileMeta(path)
		return WorkbookInfo{
			Path:        path,
			ActiveSheet: wb.Sheet,
			Rows:        rows,
			Cols:        cols,
			SheetNames:  sheets,
			ActiveCell:  active,
			FileSize:    size,
			ModTime:     mod,
			Styles:      summarizeStyles(wb),
		}, nil
	}
	if ctx.State.Workbook == nil {
		return WorkbookInfo{}, errors.New("no workbook loaded")
	}
	rows, cols := ctx.State.Workbook.MaxCoords()
	size, mod := fileMeta(ctx.State.SourcePath)
	return WorkbookInfo{
		Path:        ctx.State.SourcePath,
		ActiveSheet: ctx.State.Workbook.Sheet,
		Rows:        rows,
		Cols:        cols,
		SheetNames:  ctx.State.SheetNames,
		ActiveCell:  ctx.State.Workbook.ActiveCell,
		FileSize:    size,
		ModTime:     mod,
		Styles:      summarizeStyles(ctx.State.Workbook),
	}, nil
}

func summarizeStyles(wb *workbook.Workbook) StyleSummary {
	if wb == nil || len(wb.Styles) == 0 {
		return StyleSummary{}
	}
	summary := StyleSummary{
		FillColors:    map[string]int{},
		FontColors:    map[string]int{},
		NumberFormats: map[string]int{},
		BorderUsage:   map[string]int{},
	}
	for _, style := range wb.Styles {
		summary.StyledCells++
		if style.FillColor != "" {
			summary.FillColors[style.FillColor]++
		}
		if style.FontColor != "" {
			summary.FontColors[style.FontColor]++
		}
		if style.NumberFormat != "" {
			summary.NumberFormats[style.NumberFormat]++
		}
		if style.Bold {
			summary.BoldCells++
		}
		if style.Italic {
			summary.ItalicCells++
		}
		if style.Underline {
			summary.UnderlineCells++
		}
		if len(style.Borders) > 0 {
			for edge := range style.Borders {
				summary.BorderUsage[edge]++
			}
		}
	}
	return summary
}

func fileMeta(path string) (int64, time.Time) {
	if strings.TrimSpace(path) == "" {
		return 0, time.Time{}
	}
	info, err := os.Stat(path)
	if err != nil {
		return 0, time.Time{}
	}
	return info.Size(), info.ModTime()
}

func displayPath(path string) string {
	if strings.TrimSpace(path) == "" {
		return "(unsaved)"
	}
	abs, err := filepath.Abs(path)
	if err != nil {
		return path
	}
	return abs
}
