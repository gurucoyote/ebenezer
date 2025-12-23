package ui

import (
	"errors"
	"sync"

	"golang.org/x/term"
)

// ErrNotTerminal is returned when raw-mode operations are attempted on a non-tty.
var ErrNotTerminal = errors.New("terminal: not a terminal")

// TerminalController abstracts raw/cooked lifecycle so keyboard and MCP
// transports can share the same input handling guarantees.
type TerminalController interface {
	EnterRaw() error
	Restore() error
	WithCooked(func() error) error
}

// Terminal manages raw-mode transitions on a single file descriptor.
type Terminal struct {
	fd       int
	state    *term.State
	restore  sync.Once
	rawReady bool
	mu       sync.Mutex
}

// NewTerminal constructs a controller for the given file descriptor (usually stdin).
func NewTerminal(fd int) *Terminal {
	return &Terminal{fd: fd}
}

// EnterRaw switches the terminal into raw mode, saving the previous state for restoration.
func (t *Terminal) EnterRaw() error {
	t.mu.Lock()
	defer t.mu.Unlock()
	if t.rawReady {
		return nil
	}
	if !term.IsTerminal(t.fd) {
		return ErrNotTerminal
	}
	state, err := term.MakeRaw(t.fd)
	if err != nil {
		return err
	}
	t.state = state
	t.rawReady = true
	return nil
}

// Restore returns the terminal to its saved state.
func (t *Terminal) Restore() error {
	t.restore.Do(func() {
		t.mu.Lock()
		defer t.mu.Unlock()
		if t.state != nil {
			_ = term.Restore(t.fd, t.state)
			t.state = nil
			t.rawReady = false
		}
	})
	if t.state != nil {
		// Restore already attempted; any error would have been swallowed to avoid double prints.
		return nil
	}
	return nil
}

// WithCooked temporarily restores cooked mode for interactive prompts, then re-enters raw.
func (t *Terminal) WithCooked(fn func() error) error {
	t.mu.Lock()
	defer t.mu.Unlock()

	if !t.rawReady {
		return fn()
	}
	orig := t.state
	if err := term.Restore(t.fd, orig); err != nil {
		return err
	}
	if err := fn(); err != nil {
		// Try to re-enter raw even if fn fails.
		if rawState, rawErr := term.MakeRaw(t.fd); rawErr == nil {
			t.state = rawState
			t.rawReady = true
		}
		return err
	}
	rawState, err := term.MakeRaw(t.fd)
	if err != nil {
		t.rawReady = false
		return err
	}
	t.state = rawState
	t.rawReady = true
	return nil
}
