package cmd

import (
	"fmt"
	"strings"

	"github.com/spf13/cobra"
)

var editCmd = &cobra.Command{
	Use:   "edit [value]",
	Short: "Set the current cell's value",
	Args:  cobra.MinimumNArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		value := strings.Join(args, " ")
		appState.EditCurrentCell(value)
		fmt.Fprintf(cmd.OutOrStdout(), "%s = %q\n", appState.Address(), appState.CurrentValue())
		return nil
	},
}

var clearCmd = &cobra.Command{
	Use:   "clear",
	Short: "Clear the current cell",
	Run: func(cmd *cobra.Command, args []string) {
		appState.ClearCurrentCell()
		fmt.Fprintf(cmd.OutOrStdout(), "%s cleared\n", appState.Address())
	},
}

var yankCmd = &cobra.Command{
	Use:   "yank",
	Short: "Copy the current cell into the clipboard",
	Run: func(cmd *cobra.Command, args []string) {
		value := appState.YankCurrentCell()
		fmt.Fprintf(cmd.OutOrStdout(), "yanked %s = %q\n", appState.Address(), value)
	},
}

var cutCmd = &cobra.Command{
	Use:   "cut",
	Short: "Cut the current cell into the clipboard",
	Run: func(cmd *cobra.Command, args []string) {
		value := appState.CutCurrentCell()
		fmt.Fprintf(cmd.OutOrStdout(), "cut %s = %q\n", appState.Address(), value)
	},
}

var pasteCmd = &cobra.Command{
	Use:   "paste",
	Short: "Paste the clipboard into the current cell",
	RunE: func(cmd *cobra.Command, args []string) error {
		if err := appState.PasteClipboard(); err != nil {
			return err
		}
		fmt.Fprintf(cmd.OutOrStdout(), "pasted %q into %s\n", appState.CurrentValue(), appState.Address())
		return nil
	},
}

func init() {
	rootCmd.AddCommand(editCmd)
	rootCmd.AddCommand(clearCmd)
	rootCmd.AddCommand(yankCmd)
	rootCmd.AddCommand(cutCmd)
	rootCmd.AddCommand(pasteCmd)
}
