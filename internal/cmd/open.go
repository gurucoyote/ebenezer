package cmd

import (
	"ebenezer/internal/actions"
	"github.com/spf13/cobra"
)

var openCmd = &cobra.Command{
	Use:   "open [path]",
	Short: "Open a workbook (.csv or .xlsx)",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		params := []string{args[0]}
		if openSheet != "" {
			params = append(params, "--sheet="+openSheet)
		}
		_, err := executeAction(cmd, actions.OpenFile, params)
		return err
	},
}

var openSheet string

func init() {
	rootCmd.AddCommand(openCmd)
	openCmd.Flags().StringVar(&openSheet, "sheet", "", "Sheet name when opening .xlsx files")
}
