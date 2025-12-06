package cmd

import (
	"fmt"

	"ebenezer/internal/workbook"
	"github.com/spf13/cobra"
)

var openCmd = &cobra.Command{
	Use:   "open [path]",
	Short: "Open a CSV file into the demo workbook",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		wb, err := workbook.FromCSV(args[0])
		if err != nil {
			return err
		}
		appState.LoadWorkbook(wb)
		fmt.Fprintf(cmd.OutOrStdout(), "loaded %s (%d rows)\n", wb.Name, len(wb.Cells))
		return nil
	},
}

func init() {
	rootCmd.AddCommand(openCmd)
}
