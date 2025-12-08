package cmd

import (
	"ebenezer/internal/actions"
	"github.com/spf13/cobra"
)

var styleCmd = &cobra.Command{
	Use:   "style [address]",
	Short: "Describe the formatting of a cell (defaults to current cell)",
	Args:  cobra.MaximumNArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		_, err := executeAction(cmd, actions.StyleDescribe, args)
		return err
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
		_, err := executeAction(cmd, actions.StyleCopy, args)
		return err
	},
}

var stylePasteCmd = &cobra.Command{
	Use:   "style-paste [range]",
	Short: "Paste formatting into a cell or range",
	Args:  cobra.MaximumNArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		_, err := executeAction(cmd, actions.StylePaste, args)
		return err
	},
}
