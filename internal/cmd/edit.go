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
		selection := appState.SelectionSummary()
		appState.ClearCurrentCell()
		if selection != "" {
			fmt.Fprintf(cmd.OutOrStdout(), "cleared %s\n", selection)
			return
		}
		fmt.Fprintf(cmd.OutOrStdout(), "%s cleared\n", appState.Address())
	},
}

var yankCmd = &cobra.Command{
	Use:   "yank",
	Short: "Copy the current cell into the clipboard",
	Run: func(cmd *cobra.Command, args []string) {
		selection := appState.SelectionSummary()
		value := appState.YankCurrentCell()
		if selection != "" {
			fmt.Fprintf(cmd.OutOrStdout(), "yanked %s\n", selection)
			return
		}
		fmt.Fprintf(cmd.OutOrStdout(), "yanked %s = %q\n", appState.Address(), value)
	},
}

var cutCmd = &cobra.Command{
	Use:   "cut",
	Short: "Cut the current cell into the clipboard",
	Run: func(cmd *cobra.Command, args []string) {
		selection := appState.SelectionSummary()
		value := appState.CutCurrentCell()
		if selection != "" {
			fmt.Fprintf(cmd.OutOrStdout(), "cut %s\n", selection)
			return
		}
		fmt.Fprintf(cmd.OutOrStdout(), "cut %s = %q\n", appState.Address(), value)
	},
}

var pasteBefore bool

var pasteCmd = &cobra.Command{
	Use:   "paste",
	Short: "Paste the clipboard into the current location",
	RunE: func(cmd *cobra.Command, args []string) error {
		selection := appState.SelectionSummary()
		if err := appState.PasteClipboard(pasteBefore); err != nil {
			return err
		}
		if selection != "" {
			fmt.Fprintf(cmd.OutOrStdout(), "pasted into %s\n", selection)
			return nil
		}
		fmt.Fprintf(cmd.OutOrStdout(), "pasted into %s\n", appState.Address())
		return nil
	},
}

func init() {
	rootCmd.AddCommand(editCmd)
	rootCmd.AddCommand(clearCmd)
	rootCmd.AddCommand(yankCmd)
	rootCmd.AddCommand(cutCmd)
	rootCmd.AddCommand(pasteCmd)
	pasteCmd.Flags().BoolVar(&pasteBefore, "before", false, "Paste rows above the current row when clipboard holds rows")
}
