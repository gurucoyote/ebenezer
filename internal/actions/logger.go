package actions

import (
	"encoding/json"
	"io"
	"time"
)

// Logger captures structured information about action executions.
type Logger interface {
	Before(ctx Context, meta Metadata, args []string)
	After(ctx Context, meta Metadata, args []string, res Result, err error, duration time.Duration)
}

// NopLogger discards all events.
type NopLogger struct{}

func (NopLogger) Before(Context, Metadata, []string) {}
func (NopLogger) After(Context, Metadata, []string, Result, error, time.Duration) {}

// JSONLogger writes start/end records to an io.Writer as JSON lines.
type JSONLogger struct {
	Writer io.Writer
}

func (l JSONLogger) Before(ctx Context, meta Metadata, args []string) {
	if l.Writer == nil {
		return
	}
	event := map[string]any{
		"event":   "action_start",
		"action":  meta.Name,
		"category": meta.Category,
		"args":    args,
		"ts":      time.Now().Format(time.RFC3339Nano),
	}
	_ = json.NewEncoder(l.Writer).Encode(event)
}

func (l JSONLogger) After(ctx Context, meta Metadata, args []string, res Result, err error, dur time.Duration) {
	if l.Writer == nil {
		return
	}
	event := map[string]any{
		"event":    "action_end",
		"action":   meta.Name,
		"category": meta.Category,
		"args":     args,
		"duration_ms": float64(dur.Microseconds()) / 1000.0,
		"ts":         time.Now().Format(time.RFC3339Nano),
	}
	if err != nil {
		event["status"] = "error"
		event["error"] = err.Error()
	} else {
		event["status"] = "ok"
	}
	_ = json.NewEncoder(l.Writer).Encode(event)
}
