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
		return startKeyboardLoop(cmd)
	},
}

func init() {
	rootCmd.AddCommand(keyboardCmd)
}

func startKeyboardLoop(c *cobra.Command) error {
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
				'i': func(ctx *keyboard.Context) error {
					fmt.Fprintln(c.OutOrStdout(), "entering insert mode placeholder")
					return nil
				},
				's': func(ctx *keyboard.Context) error {
					status.Print(c.OutOrStdout(), appState)
					return nil
				},
				'g': gotoShortcut(c),
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
