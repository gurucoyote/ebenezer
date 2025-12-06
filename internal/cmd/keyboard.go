package cmd

import (
	"context"
	"fmt"

	"ebenezer/internal/ui/keyboard"
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
