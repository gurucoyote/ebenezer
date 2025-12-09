package discovery

import (
	"encoding/json"
	"testing"
)

func TestMarshalActionsProducesValidJSON(t *testing.T) {
	data, err := MarshalActions(false)
	if err != nil {
		t.Fatalf("marshal actions: %v", err)
	}
	var payload ActionList
	if err := json.Unmarshal(data, &payload); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}
	if len(payload.Actions) == 0 {
		t.Fatalf("expected at least one action in discovery payload")
	}
	if payload.Actions[0].Name == "" {
		t.Fatalf("expected action name to be populated")
	}
}
