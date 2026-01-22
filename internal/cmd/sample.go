package cmd

import (
	"ebenezer/internal/actions"
	"github.com/spf13/cobra"
)

var sampleCmd = &cobra.Command{
	Use:   "sample",
	Short: "Reload the built-in sample workbook",
	RunE: func(cmd *cobra.Command, args []string) error {
		_, err := executeAction(cmd, actions.SampleData, nil)
		return err
	},
}

func init() {
	rootCmd.AddCommand(sampleCmd)
	markInteractive(sampleCmd)
}
