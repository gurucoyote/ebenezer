package cmd

import (
	"fmt"
	"strconv"

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

var rowReorderCmd = &cobra.Command{
	Use:   "reorder",
	Short: "Reorder rows using a stable partition on a column predicate",
	RunE: func(cmd *cobra.Command, args []string) error {
		column, _ := cmd.Flags().GetString("column")
		if column == "" {
			return fmt.Errorf("required flag --column not set")
		}

		value, _ := cmd.Flags().GetString("value")
		predicate, _ := cmd.Flags().GetString("predicate")
		caseInsensitive, _ := cmd.Flags().GetBool("case-insensitive")
		separatorRows, _ := cmd.Flags().GetInt("separator-rows")
		headerRows, _ := cmd.Flags().GetInt("header-rows")
		rangeStr, _ := cmd.Flags().GetString("range")

		// Build args slice matching parseReorderArgs expectations:
		// positional: [column] [value] [predicate], then key=value pairs.
		actionArgs := []string{column}
		if value != "" {
			actionArgs = append(actionArgs, value)
		}
		if predicate != "equals_ignore_case" {
			actionArgs = append(actionArgs, predicate)
		}
		if !caseInsensitive {
			actionArgs = append(actionArgs, "case_insensitive=false")
		}
		if separatorRows != 0 {
			actionArgs = append(actionArgs, "separator_rows="+strconv.Itoa(separatorRows))
		}
		if headerRows != 1 {
			actionArgs = append(actionArgs, "header_rows="+strconv.Itoa(headerRows))
		}
		if rangeStr != "" {
			actionArgs = append(actionArgs, "range="+rangeStr)
		}

		_, err := executeAction(cmd, actions.TableReorder, actionArgs)
		return err
	},
}

func init() {
	rowReorderCmd.Flags().String("column", "", "Column index or letter (required)")
	rowReorderCmd.Flags().String("value", "", "Match value for equals/equals_ignore_case predicates")
	rowReorderCmd.Flags().String("predicate", "equals_ignore_case", "Predicate: equals|equals_ignore_case|is_blank|is_non_blank")
	rowReorderCmd.Flags().Bool("case-insensitive", true, "Make predicate case-insensitive (default true)")
	rowReorderCmd.Flags().Int("separator-rows", 0, "Blank rows inserted between partitions")
	rowReorderCmd.Flags().Int("header-rows", 1, "Number of header rows to keep fixed")
	rowReorderCmd.Flags().String("range", "", "A1-style range to restrict operation (e.g. A1:D20)")

	rowCmd.AddCommand(rowYankCmd, rowCutCmd, rowDeleteCmd, rowInsertAboveCmd, rowInsertBelowCmd, rowReorderCmd)
	rootCmd.AddCommand(rowCmd)
	markInteractive(rowCmd, rowYankCmd, rowCutCmd, rowDeleteCmd, rowInsertAboveCmd, rowInsertBelowCmd, rowReorderCmd)
}
