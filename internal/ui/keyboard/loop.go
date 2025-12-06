package keyboard

import (
	"context"
	"errors"
	"fmt"
	"io"
	"os"
	"strings"

	"github.com/chzyer/readline"
	kb "github.com/eiannone/keyboard"
)

// ErrQuit signals that the user requested to exit the keyboard loop (e.g. via 'q').
var ErrQuit = errors.New("keyboard: quit requested")

// ErrEmptyCommand is returned when a command-mode prompt is accepted without input.
var ErrEmptyCommand = errors.New("keyboard: empty command")

// CommandExecutor abstracts the command-dispatch surface. Cobra's root command
// can satisfy this via a tiny adapter so the keyboard loop can reuse existing
// CLI handlers.
type CommandExecutor interface {
	ExecuteCommand(args []string) error
}

// ExecutorFunc adapts plain functions to CommandExecutor.
type ExecutorFunc func(args []string) error

// ExecuteCommand implements CommandExecutor.
func (f ExecutorFunc) ExecuteCommand(args []string) error {
	return f(args)
}

// Action represents a callback bound to either a rune or a special key.
type Action func(*Context) error

// Bindings enumerates the callbacks used while in normal mode.
type Bindings struct {
	Runes map[rune]Action
	Keys  map[kb.Key]Action
}

// Context is passed to each binding so handlers can dispatch commands or
// examine the originating keypress.
type Context struct {
	Rune     rune
	Key      kb.Key
	Executor CommandExecutor
}

// Loop manages Vim-style normal/command modes using the shared keyboard
// driver that originated in the Gordon media controller. It keeps normal-mode
// handlers testable and reuses command-mode prompting for Ebenezer's CLI.
//
// Typical usage:
//
//	exec := keyboard.ExecutorFunc(func(args []string) error {
//	    rootCmd.SetArgs(args)
//	    return rootCmd.Execute()
//	})
//	loop := keyboard.Loop{
//	    Executor: exec,
//	    Bindings: keyboard.Bindings{
//	        Keys: map[keyboard.Key]keyboard.Action{
//	            keyboard.KeyArrowLeft:  func(ctx *keyboard.Context) error { return exec.ExecuteCommand([]string{\"move\", \"left\"}) },
//	            keyboard.KeyArrowRight: func(ctx *keyboard.Context) error { return exec.ExecuteCommand([]string{\"move\", \"right\"}) },
//	        },
//	    },
//	}
//	if err := loop.Run(context.Background()); err != nil && !errors.Is(err, keyboard.ErrQuit) {
//	    log.Fatal(err)
//	}
type Loop struct {
	Executor     CommandExecutor
	Bindings     Bindings
	Prompt       string
	QuitRunes    []rune
	CommandRunes []rune
	InfoWriter   io.Writer
}

// Run blocks until the context is cancelled, an error occurs, or a quit key
// is pressed. Callers typically invoke this from the CLI entrypoint after
// wiring bindings to command handlers.
func (l *Loop) Run(ctx context.Context) error {
	if l.Executor == nil {
		return errors.New("keyboard: executor is required")
	}
	l.ensureDefaults()

	if err := kb.Open(); err != nil {
		return fmt.Errorf("keyboard: open: %w", err)
	}
	defer kb.Close()

	for {
		select {
		case <-ctx.Done():
			return ctx.Err()
		default:
		}

		char, key, err := kb.GetKey()
		if err != nil {
			return fmt.Errorf("keyboard: read: %w", err)
		}

		if containsRune(l.CommandRunes, char) {
			if err := l.commandMode(); err != nil {
				if errors.Is(err, ErrEmptyCommand) || errors.Is(err, readline.ErrInterrupt) || errors.Is(err, io.EOF) {
					continue
				}
				return err
			}
			continue
		}

		if containsRune(l.QuitRunes, char) {
			return ErrQuit
		}

		if act := l.Bindings.Runes[char]; act != nil {
			if err := act(&Context{Rune: char, Key: key, Executor: l.Executor}); err != nil {
				if errors.Is(err, ErrQuit) {
					return err
				}
				l.reportError(err)
			}
			continue
		}

		if act := l.Bindings.Keys[key]; act != nil {
			if err := act(&Context{Rune: char, Key: key, Executor: l.Executor}); err != nil {
				if errors.Is(err, ErrQuit) {
					return err
				}
				l.reportError(err)
			}
		}
	}
}

func (l *Loop) commandMode() error {
	kb.Close()
	defer func() {
		if err := kb.Open(); err != nil {
			l.reportError(fmt.Errorf("keyboard: reopen failed: %w", err))
		}
	}()

	rl, err := readline.New(l.Prompt + " ")
	if err != nil {
		return fmt.Errorf("keyboard: prompt: %w", err)
	}
	defer rl.Close()

	line, err := rl.Readline()
	if err != nil {
		return err
	}
	if strings.TrimSpace(line) == "" {
		return ErrEmptyCommand
	}
	args := strings.Fields(line)
	return l.Executor.ExecuteCommand(args)
}

func (l *Loop) ensureDefaults() {
	if l.Prompt == "" {
		l.Prompt = ":"
	}
	if len(l.QuitRunes) == 0 {
		l.QuitRunes = []rune{'q', 'Q'}
	}
	if len(l.CommandRunes) == 0 {
		l.CommandRunes = []rune{':'}
	}
	if l.Bindings.Runes == nil {
		l.Bindings.Runes = map[rune]Action{}
	}
	if l.Bindings.Keys == nil {
		l.Bindings.Keys = map[kb.Key]Action{}
	}
	if l.InfoWriter == nil {
		l.InfoWriter = os.Stderr
	}
}

func (l *Loop) reportError(err error) {
	if err == nil {
		return
	}
	fmt.Fprintf(l.InfoWriter, "keyboard loop: %v\n", err)
}

func containsRune(set []rune, target rune) bool {
	for _, r := range set {
		if r == target {
			return true
		}
	}
	return false
}
