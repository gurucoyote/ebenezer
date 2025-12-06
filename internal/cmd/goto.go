package cmd

import (
	"fmt"
	"strings"

	"github.com/spf13/cobra"
)

var gotoCmd = &cobra.Command{
	Use:   "goto [address]",
	Short: "Move the cursor to the given cell address (e.g., B12)",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		if err := appState.Goto(args[0]); err != nil {
			return err
		}
		fmt.Fprintf(cmd.OutOrStdout(), "→ %s = %q\n", strings.ToUpper(args[0]), appState.CurrentValue())
		return nil
	},
}

var colHeaderCmd = &cobra.Command{
	Use:   "colheader",
	Short: "Print the header of the current column (row 1)",
	Run: func(cmd *cobra.Command, args []string) {
		cmd.Printf("column %d header: %q\n", appState.Cursor.Col, appState.ColumnHeader())
	},
}

var rowHeaderCmd = &cobra.Command{
	Use:   "rowheader",
	Short: "Print the header of the current row (column 1)",
	Run: func(cmd *cobra.Command, args []string) {
		cmd.Printf("row %d header: %q\n", appState.Cursor.Row, appState.RowHeader())
	},
}

func init() {
	rootCmd.AddCommand(gotoCmd)
	rootCmd.AddCommand(colHeaderCmd)
	rootCmd.AddCommand(rowHeaderCmd)
}
