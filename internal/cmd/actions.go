package cmd

import (
	"fmt"

	"ebenezer/internal/discovery"
	"github.com/spf13/cobra"
)

var (
	actionsCmd = &cobra.Command{
		Use:   "actions",
		Short: "Inspect Ebenezer action metadata",
	}

	actionsListCmd = &cobra.Command{
		Use:   "list",
		Short: "Print registered actions as JSON",
		RunE:  runActionsList,
	}
)

func init() {
	actionsCmd.AddCommand(actionsListCmd)
	rootCmd.AddCommand(actionsCmd)
}

func runActionsList(cmd *cobra.Command, _ []string) error {
	data, err := discovery.MarshalActions(true)
	if err != nil {
		return err
	}
	// Ensure CLI output ends with a newline for scripts.
	if _, err := fmt.Fprintf(cmd.OutOrStdout(), "%s\n", data); err != nil {
		return err
	}
	return nil
}
