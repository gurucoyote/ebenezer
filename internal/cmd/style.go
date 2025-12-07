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
	rootCmd.AddCommand(styleCmd, styleCopyCmd, stylePasteCmd)
}

var styleCopyCmd = &cobra.Command{
	Use:   "style-copy [range]",
	Short: "Copy formatting from a cell or range",
	Args:  cobra.MaximumNArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		target := ""
		if len(args) > 0 {
			target = args[0]
		} else {
			target = appState.Address()
		}
		if err := appState.CopyStyle(target); err != nil {
			return err
		}
		fmt.Fprintf(cmd.OutOrStdout(), "style copied from %s\n", strings.ToUpper(target))
		return nil
	},
}

var stylePasteCmd = &cobra.Command{
	Use:   "style-paste [range]",
	Short: "Paste formatting into a cell or range",
	Args:  cobra.MaximumNArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		target := ""
		if len(args) > 0 {
			target = args[0]
		} else {
			target = appState.Address()
		}
		if err := appState.PasteStyle(target); err != nil {
			return err
		}
		fmt.Fprintf(cmd.OutOrStdout(), "style pasted into %s\n", strings.ToUpper(target))
		return nil
	},
}
