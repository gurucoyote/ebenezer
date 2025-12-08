package cmd

import (
	"ebenezer/internal/actions"
	"github.com/spf13/cobra"
)

var sheetSelectCmd = &cobra.Command{
	Use:   "ps [sheet]",
	Short: "List or switch sheets in the current workbook",
	RunE: func(cmd *cobra.Command, args []string) error {
		if len(args) == 0 {
			_, err := executeAction(cmd, actions.SheetList, nil)
			return err
		}
		_, err := executeAction(cmd, actions.SheetSwitch, args)
		return err
	},
}

var sheetNewCmd = &cobra.Command{
	Use:   "ns <name> [copy-from]",
	Short: "Create a new sheet (optionally by copying an existing one)",
	Args:  cobra.MinimumNArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		copyFlag, _ := cmd.Flags().GetString("copy")
		params := append([]string{}, args...)
		if copyFlag != "" {
			params = append(params, "--copy="+copyFlag)
		}
		_, err := executeAction(cmd, actions.SheetCreate, params)
		return err
	},
}

func init() {
	sheetNewCmd.Flags().String("copy", "", "Copy contents from existing sheet")
	rootCmd.AddCommand(sheetSelectCmd, sheetNewCmd)
}
