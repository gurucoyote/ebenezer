package cmd

import (
	"fmt"
	"strings"

	"github.com/spf13/cobra"
)

var styleCmd = &cobra.Command{
	Use:   "style [address]",
	Short: "Describe the formatting of a cell (defaults to current cell)",
	Args:  cobra.MaximumNArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		addr := appState.Address()
		if len(args) == 1 {
			addr = strings.ToUpper(args[0])
		}
		style, ok := appState.StyleAt(addr)
		if !ok {
			fmt.Fprintf(cmd.OutOrStdout(), "%s: no style metadata available\n", addr)
			return nil
		}
		fmt.Fprintf(cmd.OutOrStdout(), "%s: %s\n", addr, style.Describe())
		return nil
	},
}

func init() {
	rootCmd.AddCommand(styleCmd)
}
