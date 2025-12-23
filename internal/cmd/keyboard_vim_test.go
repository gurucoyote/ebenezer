package cmd

import (
	"context"
	"errors"
	"io"
	"testing"
	"time"

	"ebenezer/internal/app"
	"ebenezer/internal/ui/keyboard"
	"ebenezer/internal/workbook"
	githubkeyboard "github.com/eiannone/keyboard"
	"github.com/spf13/cobra"
)

type fakeKey struct {
	r   rune
	key githubkeyboard.Key
}

type fakeKeyReader struct {
	events []fakeKey
	idx    int
}

func (f *fakeKeyReader) Next() (rune, githubkeyboard.Key, error) {
	if f.idx >= len(f.events) {
		return 0, 0, context.Canceled
	}
	ev := f.events[f.idx]
	f.idx++
	return ev.r, ev.key, nil
}

func runLoopWithKeys(t *testing.T, keys []fakeKey) {
	t.Helper()
	reader := &fakeKeyReader{events: keys}
	origGetKey := getKey
	defer func() { getKey = origGetKey }()
	getKey = reader.Next
	loop := keyboard.Loop{
		Executor:            keyboard.ExecutorFunc(func(args []string) error { return nil }),
		InfoWriter:          nil,
		KeyReader:           reader.Next,
		DisableKeyboardInit: true,
		Bindings: keyboard.Bindings{
			Runes: map[rune]keyboard.Action{
				'd': deleteCellShortcut(dummyCmd()),
				'y': yankShortcut(dummyCmd()),
				'x': cutShortcut(dummyCmd()),
			},
		},
	}

	ctx, cancel := context.WithTimeout(context.Background(), 2*time.Second)
	defer cancel()
	err := loop.Run(ctx)
	if err != nil && !errors.Is(err, keyboard.ErrQuit) && err != context.Canceled {
		t.Fatalf("loop error: %v", err)
	}
}

func resetState() {
	wb := workbook.SampleWorkbook()
	appState.LoadWorkbook(wb, "", []string{wb.Sheet}, "")
}

func dummyCmd() *cobra.Command {
	cmd := &cobra.Command{}
	cmd.SetOut(io.Discard)
	cmd.SetErr(io.Discard)
	return cmd
}

func TestVimDeleteRow_dd(t *testing.T) {
	resetState()
	initialRows := len(appState.Workbook.Cells)
	runLoopWithKeys(t, []fakeKey{{r: 'd'}, {r: 'd'}, {r: 'q'}})
	if got := len(appState.Workbook.Cells); got != initialRows-1 {
		t.Fatalf("expected row delete to remove 1 row, got %d rows (start %d)", got, initialRows)
	}
}

func TestVimYankRow_yy(t *testing.T) {
	resetState()
	runLoopWithKeys(t, []fakeKey{{r: 'y'}, {r: 'y'}, {r: 'q'}})
	if appState.Clipboard.Kind != app.ClipboardRow {
		t.Fatalf("expected row yank, clipboard kind %v", appState.Clipboard.Kind)
	}
	if got := len(appState.Clipboard.Rows); got != 1 {
		t.Fatalf("expected 1 row yanked, got %d", got)
	}
	if got := len(appState.Workbook.Cells); got != 5 {
		t.Fatalf("expected no rows removed on yank, got %d", got)
	}
}

func TestVimCutRow_xx(t *testing.T) {
	resetState()
	initialRows := len(appState.Workbook.Cells)
	runLoopWithKeys(t, []fakeKey{{r: 'x'}, {r: 'x'}, {r: 'q'}})
	if appState.Clipboard.Kind != app.ClipboardRow {
		t.Fatalf("expected row cut clipboard, kind %v", appState.Clipboard.Kind)
	}
	if got := len(appState.Workbook.Cells); got != initialRows-1 {
		t.Fatalf("expected row cut to remove 1 row, got %d rows", got)
	}
}
