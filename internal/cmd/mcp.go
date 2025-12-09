package cmd

import (
	"fmt"

	"ebenezer/pkg/mcp"
	"github.com/spf13/cobra"
)

var (
	mcpCmd = &cobra.Command{
		Use:   "mcp",
		Short: "Run Ebenezer in MCP server mode",
	}

	mcpServeCmd = &cobra.Command{
		Use:   "serve",
		Short: "Start the MCP server over stdio",
		RunE:  runMCPServe,
	}
)

const (
	mcpServerName    = "ebenezer"
	mcpServerVersion = "0.1.0-dev"
)

func init() {
	mcpCmd.AddCommand(mcpServeCmd)
	rootCmd.AddCommand(mcpCmd)
}

func runMCPServe(cmd *cobra.Command, _ []string) error {
	opts := mcp.Options{
		Name:    mcpServerName,
		Version: mcpServerVersion,
	}
	if err := mcp.Serve(cmd.Context(), opts); err != nil {
		return fmt.Errorf("mcp serve: %w", err)
	}
	return nil
}
