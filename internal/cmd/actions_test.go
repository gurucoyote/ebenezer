package cmd

import (
	"bytes"
	"encoding/json"
	"testing"

	"github.com/spf13/cobra"
)

func TestActionsListCommandOutputsJSON(t *testing.T) {
	out := &bytes.Buffer{}
	c := &cobra.Command{}
	c.SetOut(out)
	c.SetErr(out)
	if err := runActionsList(c, nil); err != nil {
		t.Fatalf("runActionsList: %v", err)
	}
	var payload struct {
		Actions []map[string]any `json:"actions"`
	}
	if err := json.Unmarshal(out.Bytes(), &payload); err != nil {
		t.Fatalf("unmarshal output: %v", err)
	}
	if len(payload.Actions) == 0 {
		t.Fatalf("expected discovery payload to list actions")
	}
	if _, ok := payload.Actions[0]["name"]; !ok {
		t.Fatalf("expected action entry to contain name field")
	}
}
