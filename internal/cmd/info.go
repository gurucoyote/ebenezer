package cmd

import (
	"github.com/spf13/cobra"

	"ebenezer/internal/actions"
)

var infoCmd = &cobra.Command{
	Use:   "info [file]",
	Short: "Display workbook metadata (sheets, size, cursor info)",
	Args:  cobra.MaximumNArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		details, _ := cmd.Flags().GetBool("details")
		params := append([]string{}, args...)
		if details {
			params = append(params, "--details")
		}
		_, err := executeAction(cmd, actions.Info, params)
		return err
	},
}

func init() {
	infoCmd.Flags().Bool("details", false, "Show extended metadata (file size, timestamps)")
	rootCmd.AddCommand(infoCmd)
}
