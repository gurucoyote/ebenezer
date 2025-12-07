package cmd

import (
	"fmt"
	"strings"

	"github.com/spf13/cobra"
)

var searchCmd = &cobra.Command{
	Use:   "search [pattern]",
	Short: "Search for text within the current sheet",
	Args:  cobra.MinimumNArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		pattern := strings.Join(args, " ")
		reverse, _ := cmd.Flags().GetBool("reverse")
		if err := appState.Search(pattern, !reverse); err != nil {
			return err
		}
		fmt.Fprintf(cmd.OutOrStdout(), "found %s = %q\n", appState.Address(), appState.CurrentValue())
		return nil
	},
}

var searchNextCmd = &cobra.Command{
	Use:    "search-next",
	Short:  "Repeat the last search direction (n)",
	Hidden: true,
	RunE: func(cmd *cobra.Command, args []string) error {
		if err := appState.RepeatSearch(true); err != nil {
			return err
		}
		fmt.Fprintf(cmd.OutOrStdout(), "found %s = %q\n", appState.Address(), appState.CurrentValue())
		return nil
	},
}

var searchPrevCmd = &cobra.Command{
	Use:    "search-prev",
	Short:  "Repeat the last search in the opposite direction (N)",
	Hidden: true,
	RunE: func(cmd *cobra.Command, args []string) error {
		if err := appState.RepeatSearch(false); err != nil {
			return err
		}
		fmt.Fprintf(cmd.OutOrStdout(), "found %s = %q\n", appState.Address(), appState.CurrentValue())
		return nil
	},
}

func init() {
	searchCmd.Flags().Bool("reverse", false, "search backward (same as '?')")
	rootCmd.AddCommand(searchCmd)
	rootCmd.AddCommand(searchNextCmd)
	rootCmd.AddCommand(searchPrevCmd)
}
