package cmd

import (
	"context"
	"errors"
	"fmt"
	"io"
	"strings"
	"unicode"

	"ebenezer/internal/actions"
	"ebenezer/internal/app"
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
				githubkeyboard.KeyArrowLeft:  keyboardAction(c, actions.Move, []string{"left"}),
				githubkeyboard.KeyArrowRight: keyboardAction(c, actions.Move, []string{"right"}),
				githubkeyboard.KeyArrowUp:    keyboardAction(c, actions.Move, []string{"up"}),
				githubkeyboard.KeyArrowDown:  keyboardAction(c, actions.Move, []string{"down"}),
				githubkeyboard.KeyEsc:        clearSelectionAction(c),
			},
			Runes: map[rune]keyboard.Action{
				'i': insertShortcut(c),
				'/': searchShortcut(c, false),
				'?': searchShortcut(c, true),
				'n': keyboardAction(c, actions.SearchRepeatForward, nil),
				'N': keyboardAction(c, actions.SearchRepeatBackward, nil),
				'v': visualRangeShortcut(c),
				'V': visualRowShortcut(c),
				's': keyboardAction(c, actions.Status, nil),
				'g': gotoShortcut(c),
				'c': columnHeaderShortcut(c),
				'r': rowHeaderShortcut(c),
				'y': keyboardAction(c, actions.Yank, nil),
				'Y': keyboardAction(c, actions.RowYank, nil),
				'x': keyboardAction(c, actions.Cut, nil),
				'X': keyboardAction(c, actions.RowCut, nil),
				'p': keyboardAction(c, actions.Paste, nil),
				'P': keyboardAction(c, actions.Paste, []string{"--before"}),
				'd': deleteCellShortcut(c),
				'D': keyboardAction(c, actions.RowDelete, nil),
				'O': keyboardAction(c, actions.RowInsertAbove, nil),
				'o': keyboardAction(c, actions.RowInsertBelow, nil),
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
		_, err = executeAction(c, actions.Goto, []string{addr})
		return err
	}
}

func promptForAddress(c *cobra.Command) (string, error) {
	out := c.OutOrStdout()
	current := appState.Address()
	fmt.Fprintf(out, "\nGoto cell (ESC to cancel) [%s]: ", current)
	buffer := []rune{}
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
			if len(buffer) == 0 {
				return strings.TrimSpace(strings.ToUpper(current)), nil
			}
			return strings.TrimSpace(strings.ToUpper(string(buffer))), nil
		case githubkeyboard.KeyBackspace, githubkeyboard.KeyBackspace2:
			if len(buffer) > 0 {
				buffer = buffer[:len(buffer)-1]
				fmt.Fprint(out, "\b \b")
			}
		case githubkeyboard.KeyCtrlU, githubkeyboard.KeyCtrlW:
			handleEditingControl(out, key, &buffer)
			continue
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

func columnHeaderShortcut(c *cobra.Command) keyboard.Action {
	return func(ctx *keyboard.Context) error {
		match, err := expectNextRune('t')
		if err != nil {
			return err
		}
		if !match {
			return nil
		}
		_, err = executeAction(c, actions.ColumnHeader, nil)
		return err
	}
}

func visualRangeShortcut(c *cobra.Command) keyboard.Action {
	return func(ctx *keyboard.Context) error {
		appState.ToggleSelection(app.SelectionRange)
		status.Print(c.OutOrStdout(), appState)
		return nil
	}
}

func visualRowShortcut(c *cobra.Command) keyboard.Action {
	return func(ctx *keyboard.Context) error {
		appState.ToggleSelection(app.SelectionRow)
		status.Print(c.OutOrStdout(), appState)
		return nil
	}
}

func clearSelectionAction(c *cobra.Command) keyboard.Action {
	return func(ctx *keyboard.Context) error {
		if appState.HasSelection() {
			appState.ClearSelection()
			status.Print(c.OutOrStdout(), appState)
		}
		return nil
	}
}

func keyboardAction(c *cobra.Command, action actions.Action, args []string) keyboard.Action {
	return func(ctx *keyboard.Context) error {
		_, err := executeAction(c, action, args)
		return err
	}
}

func rowHeaderShortcut(c *cobra.Command) keyboard.Action {
	return func(ctx *keyboard.Context) error {
		match, err := expectNextRune('t')
		if err != nil {
			return err
		}
		if !match {
			return nil
		}
		_, err = executeAction(c, actions.RowHeader, nil)
		return err
	}
}

func searchShortcut(c *cobra.Command, reverse bool) keyboard.Action {
	return func(ctx *keyboard.Context) error {
		initial := appState.LastSearchQuery()
		value, err := promptForSearch(c, initial, reverse)
		if err != nil {
			if errors.Is(err, errPromptCanceled) {
				return nil
			}
			return err
		}
		term := strings.TrimSpace(value)
		if term == "" {
			term = strings.TrimSpace(initial)
		}
		if term == "" {
			return nil
		}
		action := actions.SearchForward
		if reverse {
			action = actions.SearchBackward
		}
		return keyboardAction(c, action, []string{term})(ctx)
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
		return keyboardAction(c, actions.Edit, []string{value})(ctx)
	}
}

func deleteCellShortcut(c *cobra.Command) keyboard.Action {
	return func(ctx *keyboard.Context) error {
		match, err := expectNextRune('c')
		if err != nil {
			return err
		}
		if !match {
			return nil
		}
		return keyboardAction(c, actions.Clear, nil)(ctx)
	}
}

func simpleCommand(name string, args ...string) keyboard.Action {
	return func(ctx *keyboard.Context) error {
		cmdArgs := append([]string{name}, args...)
		return ctx.Executor.ExecuteCommand(cmdArgs)
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
		case githubkeyboard.KeyCtrlU, githubkeyboard.KeyCtrlW:
			handleEditingControl(out, key, &buffer)
			continue
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

func promptForSearch(c *cobra.Command, initial string, reverse bool) (string, error) {
	direction := "Forward"
	if reverse {
		direction = "Backward"
	}
	out := c.OutOrStdout()
	fmt.Fprintf(out, "\n%s search (ESC to cancel) [%s]: ", direction, initial)
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
		case githubkeyboard.KeyCtrlU, githubkeyboard.KeyCtrlW:
			handleEditingControl(out, key, &buffer)
			continue
		default:
			if isPrintable(char) {
				buffer = append(buffer, char)
				fmt.Fprint(out, string(char))
			}
		}
	}
}

func handleEditingControl(out io.Writer, key githubkeyboard.Key, buffer *[]rune) {
	switch key {
	case githubkeyboard.KeyCtrlU:
		removed := len(*buffer)
		*buffer = (*buffer)[:0]
		eraseChars(out, removed)
	case githubkeyboard.KeyCtrlW:
		removed := deleteTrailingWord(buffer)
		eraseChars(out, removed)
	}
}

func deleteTrailingWord(buffer *[]rune) int {
	b := *buffer
	if len(b) == 0 {
		return 0
	}
	i := len(b)
	for i > 0 && unicode.IsSpace(b[i-1]) {
		i--
	}
	for i > 0 && !unicode.IsSpace(b[i-1]) {
		i--
	}
	removed := len(b) - i
	*buffer = b[:i]
	return removed
}

func eraseChars(out io.Writer, count int) {
	for i := 0; i < count; i++ {
		fmt.Fprint(out, "\b \b")
	}
}
