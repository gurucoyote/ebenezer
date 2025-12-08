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
}

func (infoAction) Name() string { return "info" }

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
	}, nil
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
