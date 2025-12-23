//go:build linux || darwin

package cmd

import (
	"bytes"
	"context"
	"errors"
	"io"
	"os"
	"os/exec"
	"strings"
	"testing"
	"time"

	"github.com/creack/pty"
)

// TestKeyboardRawModePTY exercises the keyboard loop inside a pseudo-terminal to
// ensure raw-mode entry/exit and quit handling work under a controlling TTY.
func TestKeyboardRawModePTY(t *testing.T) {
	if os.Getenv("KEYBOARD_HELPER") == "1" {
		_, _ = os.Stdout.WriteString("READY\n")
		if err := runKeyboardMode(keyboardCmd); err != nil {
			// Print to stderr so the parent process can capture failures.
			_, _ = os.Stderr.WriteString(err.Error())
			os.Exit(1)
		}
		os.Exit(0)
	}

	ctx, cancel := context.WithTimeout(context.Background(), 3*time.Second)
	defer cancel()

	cmd := exec.CommandContext(ctx, os.Args[0], "-test.run=TestKeyboardRawModePTY", "--")
	cmd.Env = append(os.Environ(), "KEYBOARD_HELPER=1")

	ptmx, err := pty.Start(cmd)
	if err != nil {
		if errors.Is(err, os.ErrPermission) || strings.Contains(err.Error(), "permission denied") {
			t.Skip("skipping PTY test: permission denied in sandbox")
		}
		t.Fatalf("start pty: %v", err)
	}
	defer func() { _ = ptmx.Close() }()

	var output bytes.Buffer
	readDone := make(chan struct{})
	go func() {
		_, _ = io.Copy(&output, ptmx)
		close(readDone)
	}()

	// Wait until the helper signals readiness before sending input.
	readyCtx, readyCancel := context.WithTimeout(ctx, 2*time.Second)
	defer readyCancel()
	for {
		if strings.Contains(output.String(), "READY") {
			break
		}
		select {
		case <-readyCtx.Done():
			t.Fatalf("did not receive READY banner from helper (output so far: %q)", output.String())
		case <-time.After(10 * time.Millisecond):
		}
	}

	// Send 'q' to trigger ErrQuit and graceful shutdown message.
	if _, err := ptmx.Write([]byte("q")); err != nil {
		t.Fatalf("write quit key: %v", err)
	}

	waitErr := cmd.Wait()
	if waitErr != nil {
		t.Fatalf("keyboard helper exited with error: %v (output: %s)", waitErr, output.String())
	}

	// Close PTY to unblock reader and gather remaining output.
	_ = ptmx.Close()
	select {
	case <-readDone:
	case <-time.After(500 * time.Millisecond):
		t.Fatal("timeout waiting for PTY reader to finish")
	}

	if !strings.Contains(output.String(), "exiting keyboard mode") {
		t.Fatalf("expected quit message, got output: %q", output.String())
	}
}
