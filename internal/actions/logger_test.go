package actions

import (
	"bytes"
	"encoding/json"
	"strings"
	"testing"
	"time"
)

type fakeAction struct{}

type fakeLogger struct{
	startCalled bool
	endCalled bool
}

func (fakeAction) Name() string { return "fake" }

func (fakeAction) Metadata() Metadata {
	return Metadata{Name: "fake", Description: "fake", Category: "test"}
}

func (fakeAction) Exec(ctx Context, args []string) (Result, error) {
	return Result{Message: "ok"}, nil
}

func (l *fakeLogger) Before(Context, Metadata, []string) { l.startCalled = true }
func (l *fakeLogger) After(Context, Metadata, []string, Result, error, time.Duration) { l.endCalled = true }

func TestExecuteActionLogs(t *testing.T) {
	logger := &fakeLogger{}
	ctx := NewContext(nil, nil)
	ctx.Logger = logger
	res, err := fakeAction{}.Exec(ctx, nil)
	if err != nil { t.Fatalf("exec: %v", err) }
	if res.Message != "ok" { t.Fatalf("unexpected message") }
}

func TestJSONLoggerOutputs(t *testing.T) {
	buf := &bytes.Buffer{}
	logger := JSONLogger{Writer: buf}
	ctx := NewContext(nil, nil)
	meta := Metadata{Name: "foo", Category: "test"}
	logger.Before(ctx, meta, []string{"arg"})
	logger.After(ctx, meta, []string{"arg"}, Result{}, nil, time.Millisecond)
	lines := strings.Split(strings.TrimSpace(buf.String()), "\n")
	if len(lines) != 2 {
		t.Fatalf("expected two log lines, got %d", len(lines))
	}
	var event map[string]any
	if err := json.Unmarshal([]byte(lines[0]), &event); err != nil {
		t.Fatalf("decode start: %v", err)
	}
	action, _ := event["action"].(string)
	if action != "foo" {
		t.Fatalf("unexpected action %s", action)
	}
}
