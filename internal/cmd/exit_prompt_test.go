package cmd

import (
	"io"
	"strings"
	"testing"
)

func TestPromptExitChoice(t *testing.T) {
	cases := []struct {
		input  string
		expect exitChoice
	}{
		{"y\n", exitSave},
		{"yes\n", exitSave},
		{"n\n", exitDiscard},
		{"discard\n", exitDiscard},
		{"c\n", exitCancel},
		{"\n", exitCancel},
		{"maybe\nn\n", exitDiscard},
	}

	for _, tc := range cases {
		choice, err := promptExitChoice(strings.NewReader(tc.input), io.Discard)
		if err != nil {
			t.Fatalf("promptExitChoice error: %v", err)
		}
		if choice != tc.expect {
			t.Fatalf("expected %v, got %v for input %q", tc.expect, choice, tc.input)
		}
	}
}
