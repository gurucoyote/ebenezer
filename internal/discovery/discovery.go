package discovery

import (
	"bytes"
	"encoding/json"

	"ebenezer/internal/actions"
)

// ActionList wraps the metadata slice so downstream consumers have a stable
// JSON schema (`{"actions":[...]}`) that matches documentation and user stories.
type ActionList struct {
	Actions []actions.MetadataDTO `json:"actions"`
}

// MarshalActions returns the discovery payload encoded as JSON. When pretty is
// true the output is indented for human-facing CLI usage; otherwise it stays
// compact for machine transports such as MCP.
func MarshalActions(pretty bool) ([]byte, error) {
	payload := ActionList{Actions: actions.Discover()}
	data, err := json.Marshal(payload)
	if err != nil {
		return nil, err
	}
	if !pretty {
		return data, nil
	}
	var buf bytes.Buffer
	if err := json.Indent(&buf, data, "", "  "); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}
