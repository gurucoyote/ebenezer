package actions

import (
	"sort"
	"strings"
)

// MetadataDTO describes an action in a transport-friendly schema for MCP and
// other automated clients. The structure intentionally mirrors Metadata but
// adds JSON tags and keeps the surface stable for discovery.
type MetadataDTO struct {
	Name         string   `json:"name"`
	Description  string   `json:"description"`
	Category     string   `json:"category"`
	Args         []ArgDTO `json:"args,omitempty"`
	Idempotent   bool     `json:"idempotent"`
	Experimental bool     `json:"experimental,omitempty"`
}

// ArgDTO mirrors Arg with JSON tags for discovery payloads.
type ArgDTO struct {
	Name        string `json:"name"`
	Description string `json:"description,omitempty"`
	Optional    bool   `json:"optional,omitempty"`
	Variadic    bool   `json:"variadic,omitempty"`
}

// Discover returns action metadata serialized into a deterministic, JSON-ready
// schema suitable for MCP discovery endpoints.
func Discover() []MetadataDTO {
	metas := ListMetadata()
	sort.Slice(metas, func(i, j int) bool {
		return strings.ToLower(metas[i].Name) < strings.ToLower(metas[j].Name)
	})
	result := make([]MetadataDTO, 0, len(metas))
	for _, meta := range metas {
		result = append(result, convertMetadata(meta))
	}
	return result
}

func convertMetadata(meta Metadata) MetadataDTO {
	args := make([]ArgDTO, len(meta.Args))
	for i, arg := range meta.Args {
		args[i] = ArgDTO{
			Name:        arg.Name,
			Description: arg.Description,
			Optional:    arg.Optional,
			Variadic:    arg.Variadic,
		}
	}
	return MetadataDTO{
		Name:         meta.Name,
		Description:  meta.Description,
		Category:     meta.Category,
		Args:         args,
		Idempotent:   meta.Idempotent,
		Experimental: meta.Experimental,
	}
}
