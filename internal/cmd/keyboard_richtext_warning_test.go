package cmd

import (
	"testing"

	"ebenezer/internal/app"
)

func TestWarnOnRichTextYankCell(t *testing.T) {
	resetState()
	appState.Workbook.RichTextRuns = map[string]int{"A1": 2}
	appState.Workbook.RichTextSheet = appState.Workbook.Sheet

	runLoopWithKeys(t, []fakeKey{
		{r: 'y'},
		{r: 'a'},
		{r: 'n'},
		{r: 'q'},
	})

	if appState.Clipboard.Kind != app.ClipboardNone {
		t.Fatalf("expected clipboard untouched when canceling rich text warning, got %v", appState.Clipboard.Kind)
	}
}
