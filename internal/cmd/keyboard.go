package cmd

import (
	"context"
	"errors"
	"fmt"
	"strings"
	"unicode"

	"ebenezer/internal/ui/keyboard"
	"ebenezer/internal/ui/status"
	githubkeyboard "github.com/eiannone/keyboard"
	"github.com/spf13/cobra"
)

var keyboardCmd = &cobra.Command{
	Use:   "keyboard",
	Short: "Enter vim-like keyboard control mode",
	RunE: func(cmd *cobra.Command, args []string) error {
		return runKeyboardMode(cmd)
	},
}

func init() {
	rootCmd.AddCommand(keyboardCmd)
}

func runKeyboardMode(c *cobra.Command) error {
	ctx := c.Context()
	if ctx == nil {
		ctx = context.Background()
	}

	exec := keyboard.ExecutorFunc(func(args []string) error {
		rootCmd.SetArgs(args)
		return rootCmd.ExecuteContext(ctx)
	})

	loop := keyboard.Loop{
		Executor:   exec,
		InfoWriter: c.ErrOrStderr(),
		Bindings: keyboard.Bindings{
			Keys: map[githubkeyboard.Key]keyboard.Action{
				githubkeyboard.KeyArrowLeft:  moveAction("left"),
				githubkeyboard.KeyArrowRight: moveAction("right"),
				githubkeyboard.KeyArrowUp:    moveAction("up"),
				githubkeyboard.KeyArrowDown:  moveAction("down"),
			},
			Runes: map[rune]keyboard.Action{
				'i': insertShortcut(c),
				's': func(ctx *keyboard.Context) error {
					status.Print(c.OutOrStdout(), appState)
					return nil
				},
				'g': gotoShortcut(c),
				'c': columnHeaderShortcut(),
				'r': rowHeaderShortcut(),
				'y': simpleCommand("yank"),
				'x': simpleCommand("cut"),
				'p': simpleCommand("paste"),
				'd': deleteCellShortcut(),
			},
		},
	}

	if err := loop.Run(ctx); err != nil {
		if err == keyboard.ErrQuit {
			fmt.Fprintln(c.OutOrStdout(), "exiting keyboard mode")
			return nil
		}
		return err
	}
	return nil
}

func moveAction(direction string) keyboard.Action {
	return func(ctx *keyboard.Context) error {
		return ctx.Executor.ExecuteCommand([]string{"move", direction})
	}
}

var errPromptCanceled = errors.New("prompt cancelled")

func gotoShortcut(c *cobra.Command) keyboard.Action {
	return func(ctx *keyboard.Context) error {
		addr, err := promptForAddress(c)
		if err != nil {
			if errors.Is(err, errPromptCanceled) {
				return nil
			}
			return err
		}
		if addr == "" {
			return nil
		}
		return ctx.Executor.ExecuteCommand([]string{"goto", addr})
	}
}

func promptForAddress(c *cobra.Command) (string, error) {
	out := c.OutOrStdout()
	current := appState.Address()
	fmt.Fprintf(out, "\nGoto cell (ESC to cancel) [%s]: ", current)
	buffer := []rune(current)
	fmt.Fprint(out, current)
	for {
		char, key, err := githubkeyboard.GetKey()
		if err != nil {
			return "", err
		}
		switch key {
		case githubkeyboard.KeyEsc:
			fmt.Fprintln(out)
			return "", errPromptCanceled
		case githubkeyboard.KeyEnter:
			fmt.Fprintln(out)
			return strings.TrimSpace(strings.ToUpper(string(buffer))), nil
		case githubkeyboard.KeyBackspace, githubkeyboard.KeyBackspace2:
			if len(buffer) > 0 {
				buffer = buffer[:len(buffer)-1]
				fmt.Fprint(out, "\b \b")
			}
		default:
			if unicode.IsLetter(char) {
				char = unicode.ToUpper(char)
				buffer = append(buffer, char)
				fmt.Fprint(out, string(char))
			} else if unicode.IsDigit(char) {
				buffer = append(buffer, char)
				fmt.Fprint(out, string(char))
			}
		}
	}
}

func columnHeaderShortcut() keyboard.Action {
	return func(ctx *keyboard.Context) error {
		match, err := expectNextRune('t')
		if err != nil {
			return err
		}
		if !match {
			return nil
		}
		return ctx.Executor.ExecuteCommand([]string{"colheader"})
	}
}

func rowHeaderShortcut() keyboard.Action {
	return func(ctx *keyboard.Context) error {
		match, err := expectNextRune('t')
		if err != nil {
			return err
		}
		if !match {
			return nil
		}
		return ctx.Executor.ExecuteCommand([]string{"rowheader"})
	}
}

func expectNextRune(target rune) (bool, error) {
	char, _, err := githubkeyboard.GetKey()
	if err != nil {
		return false, err
	}
	return unicode.ToLower(char) == unicode.ToLower(target), nil
}

func insertShortcut(c *cobra.Command) keyboard.Action {
	return func(ctx *keyboard.Context) error {
		value, err := promptForText(c, appState.CurrentValue())
		if err != nil {
			if errors.Is(err, errPromptCanceled) {
				return nil
			}
			return err
		}
		return ctx.Executor.ExecuteCommand([]string{"edit", value})
	}
}

func deleteCellShortcut() keyboard.Action {
	return func(ctx *keyboard.Context) error {
		match, err := expectNextRune('c')
		if err != nil {
			return err
		}
		if !match {
			return nil
		}
		return ctx.Executor.ExecuteCommand([]string{"clear"})
	}
}

func simpleCommand(name string) keyboard.Action {
	return func(ctx *keyboard.Context) error {
		return ctx.Executor.ExecuteCommand([]string{name})
	}
}

func promptForText(c *cobra.Command, initial string) (string, error) {
	out := c.OutOrStdout()
	fmt.Fprintf(out, "\nEnter value (ESC to cancel) [%s]: ", initial)
	buffer := []rune(initial)
	fmt.Fprint(out, initial)
	for {
		char, key, err := githubkeyboard.GetKey()
		if err != nil {
			return "", err
		}
		switch key {
		case githubkeyboard.KeyEsc:
			fmt.Fprintln(out)
			return "", errPromptCanceled
		case githubkeyboard.KeyEnter:
			fmt.Fprintln(out)
			return string(buffer), nil
		case githubkeyboard.KeyBackspace, githubkeyboard.KeyBackspace2:
			if len(buffer) > 0 {
				buffer = buffer[:len(buffer)-1]
				fmt.Fprint(out, "\b \b")
			}
		default:
			if isPrintable(char) {
				buffer = append(buffer, char)
				fmt.Fprint(out, string(char))
			}
		}
	}
}

func isPrintable(r rune) bool {
	return r >= 32 && r != 127
}
