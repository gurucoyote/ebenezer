package cmd

import (
	"errors"
	"io"
	"testing"
	"time"

	"ebenezer/internal/actions"
	"ebenezer/internal/ui/keyboard"
	githubkeyboard "github.com/eiannone/keyboard"
)

type stubExecutor struct {
	calls [][]string
	err   error
}

func (s *stubExecutor) ExecuteCommand(args []string) error {
	s.calls = append(s.calls, append([]string(nil), args...))
	return s.err
}

func TestKeyReaderWithCtrlArrows_CtrlLeftDispatchesMoveSpan(t *testing.T) {
	orig := getKey
	origPoll := pollReadable
	defer func() { getKey = orig }()
	defer func() { pollReadable = origPoll }()

	events := []keyboardEvent{
		{r: 0, k: githubkeyboard.KeyEsc},
		{r: '[', k: 0},
		{r: '1', k: 0},
		{r: ';', k: 0},
		{r: '5', k: 0},
		{r: 'D', k: 0},
	}
	getKey = func() (rune, githubkeyboard.Key, error) {
		if len(events) == 0 {
			return 0, 0, io.EOF
		}
		ev := events[0]
		events = events[1:]
		return ev.r, ev.k, nil
	}
	pollReadable = func(timeout time.Duration) (bool, error) { return true, nil }

	exec := &stubExecutor{}
	reader := keyReaderWithCtrlArrows(exec)

	r, k, err := reader()
	if err != nil {
		t.Fatalf("reader: %v", err)
	}
	if r != 0 || k != 0 {
		t.Fatalf("expected empty keypress, got r=%q k=%v", r, k)
	}
	if got, want := len(exec.calls), 1; got != want {
		t.Fatalf("expected %d exec calls, got %d", want, got)
	}
	if got, want := exec.calls[0], []string{"move-span", "left"}; !equalStrings(got, want) {
		t.Fatalf("expected exec call %v, got %v", want, got)
	}
}

func TestCtrlArrowDirectionVariants(t *testing.T) {
	cases := []struct {
		seq string
		dir string
		ok  bool
	}{
		{seq: "[1;5D", dir: "left", ok: true},
		{seq: "[1;5C", dir: "right", ok: true},
		{seq: "[1;5A", dir: "up", ok: true},
		{seq: "[1;5B", dir: "down", ok: true},
		{seq: "[5D", dir: "left", ok: true},
		{seq: "5C", dir: "right", ok: true},
		{seq: "[1;2C", dir: "", ok: false},
		{seq: "[D", dir: "", ok: false},
		{seq: "", dir: "", ok: false},
	}

	for _, tc := range cases {
		t.Run(tc.seq, func(t *testing.T) {
			dir, ok := ctrlArrowDirection(tc.seq)
			if ok != tc.ok {
				t.Fatalf("expected ok=%v, got %v", tc.ok, ok)
			}
			if ok && dir != tc.dir {
				t.Fatalf("expected dir=%q, got %q", tc.dir, dir)
			}
		})
	}
}

type actionExecutor struct {
	err error
}

func (a *actionExecutor) ExecuteCommand(args []string) error {
	if len(args) == 0 {
		return nil
	}
	switch args[0] {
	case "move-span":
		if len(args) < 2 {
			return errors.New("move-span requires direction")
		}
		ctx := actions.NewContext(appState, io.Discard)
		ctx.Logger = actions.NopLogger{}
		_, err := actions.MoveSpan.Exec(ctx, args[1:])
		return err
	default:
		return a.err
	}
}

func TestKeyReaderWithCtrlArrows_UpdatesState(t *testing.T) {
	resetState()
	appState.Cursor.Row = 1
	appState.Cursor.Col = 1

	orig := getKey
	origPoll := pollReadable
	defer func() { getKey = orig }()
	defer func() { pollReadable = origPoll }()

	events := []keyboardEvent{
		{r: 0, k: githubkeyboard.KeyEsc},
		{r: '[', k: 0},
		{r: '5', k: 0},
		{r: 'C', k: 0},
	}
	getKey = func() (rune, githubkeyboard.Key, error) {
		if len(events) == 0 {
			return 0, 0, io.EOF
		}
		ev := events[0]
		events = events[1:]
		return ev.r, ev.k, nil
	}
	pollReadable = func(timeout time.Duration) (bool, error) { return true, nil }

	exec := &actionExecutor{}
	reader := keyReaderWithCtrlArrows(exec)
	if _, _, err := reader(); err != nil {
		t.Fatalf("reader: %v", err)
	}
	if got := appState.Address(); got != "C1" {
		t.Fatalf("expected cursor to move to C1, got %s", got)
	}
}

func TestKeyReaderWithCtrlArrows_EscFollowedByRuneDoesNotEatNextKey(t *testing.T) {
	orig := getKey
	origPoll := pollReadable
	defer func() { getKey = orig }()
	defer func() { pollReadable = origPoll }()

	events := []keyboardEvent{
		{r: 0, k: githubkeyboard.KeyEsc},
		{r: 'x', k: 0},
	}
	getKey = func() (rune, githubkeyboard.Key, error) {
		if len(events) == 0 {
			return 0, 0, io.EOF
		}
		ev := events[0]
		events = events[1:]
		return ev.r, ev.k, nil
	}
	pollReadable = func(timeout time.Duration) (bool, error) { return true, nil }

	exec := &stubExecutor{}
	reader := keyReaderWithCtrlArrows(exec)

	r, k, err := reader()
	if err != nil {
		t.Fatalf("reader esc: %v", err)
	}
	if k != githubkeyboard.KeyEsc {
		t.Fatalf("expected esc key, got %v (r=%q)", k, r)
	}

	r, k, err = reader()
	if err != nil {
		t.Fatalf("reader next: %v", err)
	}
	if r != 'x' || k != 0 {
		t.Fatalf("expected buffered 'x', got r=%q k=%v", r, k)
	}
	if len(exec.calls) != 0 {
		t.Fatalf("expected no exec calls, got %v", exec.calls)
	}
}

func TestKeyReaderWithCtrlArrows_UnknownEscapeSequenceIsReplayed(t *testing.T) {
	orig := getKey
	origPoll := pollReadable
	defer func() { getKey = orig }()
	defer func() { pollReadable = origPoll }()

	events := []keyboardEvent{
		{r: 0, k: githubkeyboard.KeyEsc},
		{r: '[', k: 0},
		{r: '9', k: 0},
		{r: '9', k: 0},
		{r: 'Z', k: 0},
	}
	getKey = func() (rune, githubkeyboard.Key, error) {
		if len(events) == 0 {
			return 0, 0, io.EOF
		}
		ev := events[0]
		events = events[1:]
		return ev.r, ev.k, nil
	}
	pollReadable = func(timeout time.Duration) (bool, error) { return true, nil }

	exec := &stubExecutor{}
	reader := keyReaderWithCtrlArrows(exec)

	// First read returns Esc (sequence is unknown so it should be replayed after).
	_, k, err := reader()
	if err != nil {
		t.Fatalf("reader esc: %v", err)
	}
	if k != githubkeyboard.KeyEsc {
		t.Fatalf("expected esc key, got %v", k)
	}

	var replayed []rune
	for i := 0; i < 4; i++ {
		r, _, err := reader()
		if err != nil {
			t.Fatalf("reader replay %d: %v", i, err)
		}
		replayed = append(replayed, r)
	}
	if got, want := string(replayed), "[99Z"; got != want {
		t.Fatalf("expected replay %q, got %q", want, got)
	}
	if len(exec.calls) != 0 {
		t.Fatalf("expected no exec calls, got %v", exec.calls)
	}
}

func TestKeyReaderWithCtrlArrows_ImplementsKeyboardKeyReaderContract(t *testing.T) {
	orig := getKey
	origPoll := pollReadable
	defer func() { getKey = orig }()
	defer func() { pollReadable = origPoll }()

	getKey = func() (rune, githubkeyboard.Key, error) {
		return 0, 0, io.EOF
	}
	pollReadable = func(timeout time.Duration) (bool, error) { return true, nil }

	exec := keyboard.ExecutorFunc(func(args []string) error { return nil })
	reader := keyReaderWithCtrlArrows(exec)

	_, _, err := reader()
	if !errors.Is(err, io.EOF) {
		t.Fatalf("expected EOF, got %v", err)
	}
}

func equalStrings(a, b []string) bool {
	if len(a) != len(b) {
		return false
	}
	for i := range a {
		if a[i] != b[i] {
			return false
		}
	}
	return true
}
