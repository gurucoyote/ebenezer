package cmd

import (
	"github.com/spf13/cobra"

	"ebenezer/internal/actions"
)

var gotoCmd = &cobra.Command{
	Use:   "goto [address]",
	Short: "Move the cursor to the given cell address (e.g., B12)",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		_, err := executeAction(cmd, actions.Goto, args)
		return err
	},
}

var colHeaderCmd = &cobra.Command{
	Use:   "colheader",
	Short: "Print the header of the current column (row 1)",
	RunE: func(cmd *cobra.Command, args []string) error {
		_, err := executeAction(cmd, actions.ColumnHeader, nil)
		return err
	},
}

var rowHeaderCmd = &cobra.Command{
	Use:   "rowheader",
	Short: "Print the header of the current row (column 1)",
	RunE: func(cmd *cobra.Command, args []string) error {
		_, err := executeAction(cmd, actions.RowHeader, nil)
		return err
	},
}

func init() {
	rootCmd.AddCommand(gotoCmd)
	rootCmd.AddCommand(colHeaderCmd)
	rootCmd.AddCommand(rowHeaderCmd)
}
