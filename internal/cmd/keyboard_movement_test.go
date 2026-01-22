package cmd

import (
	"context"
	"errors"
	"testing"
	"time"

	"ebenezer/internal/actions"
	"ebenezer/internal/ui/keyboard"
	githubkeyboard "github.com/eiannone/keyboard"
)

func runLoopWithBindings(t *testing.T, keys []fakeKey, bindings keyboard.Bindings) {
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
		Bindings:            bindings,
	}

	ctx, cancel := context.WithTimeout(context.Background(), 2*time.Second)
	defer cancel()
	err := loop.Run(ctx)
	if err != nil && !errors.Is(err, keyboard.ErrQuit) && err != context.Canceled {
		t.Fatalf("loop error: %v", err)
	}
}

func TestKeyboardArrowRightMovesCursor(t *testing.T) {
	resetState()

	bindings := keyboard.Bindings{
		Keys: map[githubkeyboard.Key]keyboard.Action{
			githubkeyboard.KeyArrowRight: keyboardAction(dummyCmd(), actions.Move, []string{"right"}),
		},
	}

	runLoopWithBindings(t, []fakeKey{{key: githubkeyboard.KeyArrowRight}, {r: 'q'}}, bindings)

	if got := appState.Address(); got != "B1" {
		t.Fatalf("expected cursor at B1, got %s", got)
	}
}

func TestKeyboardHJKLMovesCursor(t *testing.T) {
	resetState()
	appState.Cursor.Row = 2
	appState.Cursor.Col = 2
	appState.Workbook.ActiveCell = "B2"

	bindings := keyboard.Bindings{
		Runes: map[rune]keyboard.Action{
			'h': keyboardAction(dummyCmd(), actions.Move, []string{"left"}),
			'j': keyboardAction(dummyCmd(), actions.Move, []string{"down"}),
			'k': keyboardAction(dummyCmd(), actions.Move, []string{"up"}),
			'l': keyboardAction(dummyCmd(), actions.Move, []string{"right"}),
		},
	}

	runLoopWithBindings(t, []fakeKey{
		{r: 'h'},
		{r: 'j'},
		{r: 'k'},
		{r: 'l'},
		{r: 'q'},
	}, bindings)

	if got := appState.Address(); got != "B2" {
		t.Fatalf("expected cursor at B2, got %s", got)
	}
}

func TestKeyboardGotoSpanWithArrow(t *testing.T) {
	resetState()

	bindings := keyboard.Bindings{
		Runes: map[rune]keyboard.Action{
			'g': gotoShortcut(dummyCmd()),
		},
	}

	runLoopWithBindings(t, []fakeKey{
		{r: 'g'},
		{key: githubkeyboard.KeyArrowDown},
		{r: 'q'},
	}, bindings)

	if got := appState.Address(); got != "A5" {
		t.Fatalf("expected cursor at A5, got %s", got)
	}
}

func TestKeyboardGotoSpanWithHjkl(t *testing.T) {
	resetState()

	bindings := keyboard.Bindings{
		Runes: map[rune]keyboard.Action{
			'g': gotoShortcut(dummyCmd()),
		},
	}

	runLoopWithBindings(t, []fakeKey{
		{r: 'g'},
		{r: 'l'},
		{r: 'q'},
	}, bindings)

	if got := appState.Address(); got != "C1" {
		t.Fatalf("expected cursor at C1, got %s", got)
	}
}
