package cmd

import (
	"ebenezer/internal/actions"
	"github.com/spf13/cobra"
)

var rowCmd = &cobra.Command{
	Use:   "row",
	Short: "Row operations",
}

var rowYankCmd = &cobra.Command{
	Use:   "yank",
	Short: "Yank the current row into the clipboard",
	RunE: func(cmd *cobra.Command, args []string) error {
		_, err := executeAction(cmd, actions.RowYank, nil)
		return err
	},
}

var rowCutCmd = &cobra.Command{
	Use:   "cut",
	Short: "Cut the current row into the clipboard",
	RunE: func(cmd *cobra.Command, args []string) error {
		_, err := executeAction(cmd, actions.RowCut, nil)
		return err
	},
}

var rowDeleteCmd = &cobra.Command{
	Use:   "delete",
	Short: "Delete the current row",
	RunE: func(cmd *cobra.Command, args []string) error {
		_, err := executeAction(cmd, actions.RowDelete, nil)
		return err
	},
}

var rowInsertAboveCmd = &cobra.Command{
	Use:   "insert-above",
	Short: "Insert a blank row above the cursor",
	RunE: func(cmd *cobra.Command, args []string) error {
		_, err := executeAction(cmd, actions.RowInsertAbove, nil)
		return err
	},
}

var rowInsertBelowCmd = &cobra.Command{
	Use:   "insert-below",
	Short: "Insert a blank row below the cursor",
	RunE: func(cmd *cobra.Command, args []string) error {
		_, err := executeAction(cmd, actions.RowInsertBelow, nil)
		return err
	},
}

func init() {
	rowCmd.AddCommand(rowYankCmd, rowCutCmd, rowDeleteCmd, rowInsertAboveCmd, rowInsertBelowCmd)
	rootCmd.AddCommand(rowCmd)
	markInteractive(rowCmd, rowYankCmd, rowCutCmd, rowDeleteCmd, rowInsertAboveCmd, rowInsertBelowCmd)
}
