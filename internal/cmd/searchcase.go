package cmd

import (
	"errors"
	"fmt"
	"strings"

	"github.com/spf13/cobra"
)

var searchCaseCmd = &cobra.Command{
	Use:   "search-case [sensitive|insensitive|toggle]",
	Short: "View or change the search case-sensitivity setting",
	Args:  cobra.MaximumNArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		if len(args) == 0 {
			printSearchCase(cmd)
			return nil
		}
		mode := strings.ToLower(args[0])
		switch mode {
		case "sensitive", "on", "true", "1":
			appState.SetSearchCaseSensitivity(true)
		case "insensitive", "off", "false", "0":
			appState.SetSearchCaseSensitivity(false)
		case "toggle":
			appState.SetSearchCaseSensitivity(!appState.SearchCaseSensitive)
		default:
			return errors.New("expected sensitive|insensitive|toggle")
		}
		printSearchCase(cmd)
		return nil
	},
}

func printSearchCase(cmd *cobra.Command) {
	mode := "insensitive"
	if appState.SearchCaseSensitive {
		mode = "sensitive"
	}
	fmt.Fprintf(cmd.OutOrStdout(), "search is %s\n", mode)
}

func init() {
	rootCmd.AddCommand(searchCaseCmd)
}
