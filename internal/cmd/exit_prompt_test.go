package cmd

import (
	"os"
	"testing"
)

func TestWriteQuitCancelDoesNotExit(t *testing.T) {
	resetState()
	appState.EditCurrentCell("dirty")

	tmp, err := os.CreateTemp("", "ebz-exit-*")
	if err != nil {
		t.Fatalf("temp file: %v", err)
	}
	defer func() { _ = os.Remove(tmp.Name()) }()
	if _, err := tmp.WriteString("\n"); err != nil {
		t.Fatalf("write temp: %v", err)
	}
	if _, err := tmp.Seek(0, 0); err != nil {
		t.Fatalf("seek temp: %v", err)
	}
	stdin := os.Stdin
	defer func() { os.Stdin = stdin }()
	os.Stdin = tmp

	cmd := dummyCmd()
	if err := writeQuitCmd(cmd).RunE(cmd, nil); err != nil {
		t.Fatalf("expected no error, got %v", err)
	}
}
