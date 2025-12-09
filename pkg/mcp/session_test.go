package mcp

import (
	"path/filepath"
	"testing"

	"ebenezer/internal/workbook"
)

func TestSessionManagerOpenClose(t *testing.T) {
	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "sample.csv")
	if err := workbook.SampleWorkbook().Save(path); err != nil {
		t.Fatalf("save sample: %v", err)
	}
	mgr := NewSessionManager()
	session, err := mgr.Open(path, "", false)
	if err != nil {
		t.Fatalf("open session: %v", err)
	}
	if session.ID == "" {
		t.Fatalf("expected session id")
	}
	if session.State == nil || session.State.Workbook == nil {
		t.Fatalf("expected workbook loaded")
	}
	if _, err := mgr.Get(session.ID); err != nil {
		t.Fatalf("get session: %v", err)
	}
	if !mgr.Close(session.ID) {
		t.Fatalf("expected close to succeed")
	}
	if _, err := mgr.Get(session.ID); err == nil {
		t.Fatalf("expected error after close")
	}
}
