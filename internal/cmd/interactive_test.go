package cmd

import (
	"testing"

	"github.com/spf13/cobra"
)

func TestSetInteractiveCommandVisibility(t *testing.T) {
	root := &cobra.Command{Use: "root"}
	interactive := &cobra.Command{
		Use: "edit",
		Annotations: map[string]string{
			interactiveAnnotation: "true",
		},
	}
	nonInteractive := &cobra.Command{Use: "open"}
	root.AddCommand(interactive)
	root.AddCommand(nonInteractive)

	setInteractiveCommandVisibility(root, false)
	if !interactive.Hidden {
		t.Fatalf("expected interactive command to be hidden")
	}
	if nonInteractive.Hidden {
		t.Fatalf("expected non-interactive command to remain visible")
	}

	setInteractiveCommandVisibility(root, true)
	if interactive.Hidden {
		t.Fatalf("expected interactive command to be visible")
	}
	if nonInteractive.Hidden {
		t.Fatalf("expected non-interactive command to remain visible")
	}
}
