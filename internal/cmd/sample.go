package cmd

import (
	"fmt"

	"ebenezer/internal/workbook"
	"github.com/spf13/cobra"
)

var sampleCmd = &cobra.Command{
	Use:   "sample",
	Short: "Reload the built-in sample workbook",
	Run: func(cmd *cobra.Command, args []string) {
		appState.LoadWorkbook(workbook.SampleWorkbook())
		fmt.Fprintln(cmd.OutOrStdout(), "sample workbook loaded")
	},
}

func init() {
	rootCmd.AddCommand(sampleCmd)
}
