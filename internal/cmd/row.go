package cmd

import (
	"fmt"

	"github.com/spf13/cobra"
)

var rowCmd = &cobra.Command{
	Use:   "row",
	Short: "Row operations",
}

var rowYankCmd = &cobra.Command{
	Use:   "yank",
	Short: "Yank the current row into the clipboard",
	Run: func(cmd *cobra.Command, args []string) {
		appState.YankCurrentRow()
		fmt.Fprintf(cmd.OutOrStdout(), "yanked row %d\n", appState.Cursor.Row)
	},
}

var rowCutCmd = &cobra.Command{
	Use:   "cut",
	Short: "Cut the current row into the clipboard",
	Run: func(cmd *cobra.Command, args []string) {
		appState.CutCurrentRow()
		fmt.Fprintf(cmd.OutOrStdout(), "cut row %d\n", appState.Cursor.Row)
	},
}

var rowDeleteCmd = &cobra.Command{
	Use:   "delete",
	Short: "Delete the current row",
	Run: func(cmd *cobra.Command, args []string) {
		appState.DeleteCurrentRow()
		fmt.Fprintln(cmd.OutOrStdout(), "row deleted")
	},
}

var rowInsertAboveCmd = &cobra.Command{
	Use:   "insert-above",
	Short: "Insert a blank row above the cursor",
	Run: func(cmd *cobra.Command, args []string) {
		appState.InsertRowAbove()
		fmt.Fprintf(cmd.OutOrStdout(), "inserted row above %d\n", appState.Cursor.Row)
	},
}

var rowInsertBelowCmd = &cobra.Command{
	Use:   "insert-below",
	Short: "Insert a blank row below the cursor",
	Run: func(cmd *cobra.Command, args []string) {
		appState.InsertRowBelow()
		fmt.Fprintf(cmd.OutOrStdout(), "inserted row below %d\n", appState.Cursor.Row)
	},
}

func init() {
	rowCmd.AddCommand(rowYankCmd, rowCutCmd, rowDeleteCmd, rowInsertAboveCmd, rowInsertBelowCmd)
	rootCmd.AddCommand(rowCmd)
}
