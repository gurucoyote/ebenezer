package mcp

import (
	"crypto/rand"
	"encoding/hex"
	"fmt"
	"sync"
	"time"

	"ebenezer/internal/app"
	"ebenezer/internal/workbook"
)

// Session represents an MCP-controlled workbook editing context.
type Session struct {
	ID        string
	State     *app.State
	ReadOnly  bool
	OpenedAt  time.Time
	sourceRef string
}

// SessionManager tracks active MCP sessions.
type SessionManager struct {
	mu       sync.RWMutex
	sessions map[string]*Session
}

// NewSessionManager creates an empty manager.
func NewSessionManager() *SessionManager {
	return &SessionManager{
		sessions: make(map[string]*Session),
	}
}

// Open loads a workbook from disk and returns a managed session.
func (m *SessionManager) Open(path, sheet string, readOnly bool) (*Session, error) {
	wb, sheets, active, err := workbook.FromFile(path, sheet)
	if err != nil {
		return nil, err
	}
	state := &app.State{}
	state.LoadWorkbook(wb, path, sheets, active)
	session := &Session{
		ID:        newSessionID(),
		State:     state,
		ReadOnly:  readOnly,
		OpenedAt:  time.Now().UTC(),
		sourceRef: path,
	}
	m.mu.Lock()
	m.sessions[session.ID] = session
	m.mu.Unlock()
	return session, nil
}

// Close removes a session. Returns true if it existed.
func (m *SessionManager) Close(id string) bool {
	m.mu.Lock()
	defer m.mu.Unlock()
	if _, ok := m.sessions[id]; !ok {
		return false
	}
	delete(m.sessions, id)
	return true
}

// Get returns the session by ID.
func (m *SessionManager) Get(id string) (*Session, error) {
	m.mu.RLock()
	defer m.mu.RUnlock()
	s, ok := m.sessions[id]
	if !ok {
		return nil, fmt.Errorf("session %s not found", id)
	}
	return s, nil
}

func newSessionID() string {
	var buf [16]byte
	if _, err := rand.Read(buf[:]); err == nil {
		return hex.EncodeToString(buf[:])
	}
	return fmt.Sprintf("session-%d", time.Now().UnixNano())
}
