package cmd

import (
	"context"
	"errors"
	"fmt"
	"io"
	"os"
	"os/signal"
	"strconv"
	"strings"
	"time"
	"unicode"

	"ebenezer/internal/actions"
	"ebenezer/internal/app"
	"ebenezer/internal/ui"
	"ebenezer/internal/ui/keyboard"
	"ebenezer/internal/ui/status"
	githubkeyboard "github.com/eiannone/keyboard"
	"github.com/spf13/cobra"
	"golang.org/x/sys/unix"
)

var keyboardCmd = &cobra.Command{
	Use:   "keyboard",
	Short: "Enter vim-like keyboard control mode",
	RunE: func(cmd *cobra.Command, args []string) error {
		return runKeyboardMode(cmd)
	},
}

func init() {
	rootCmd.AddCommand(keyboardCmd)
}

func runKeyboardMode(c *cobra.Command) error {
	ctx := c.Context()
	if ctx == nil {
		ctx = context.Background()
	}
	ctx, cancel := context.WithCancel(ctx)
	defer cancel()
	suppressActionLogs = true
	defer func() { suppressActionLogs = false }()
	setInteractiveCommandVisibility(rootCmd, true)
	defer setInteractiveCommandVisibility(rootCmd, false)
	registerQuitCommands(c)
	defer unregisterQuitCommands()
	registerQuitCommands(c)

	sigCh := make(chan os.Signal, 1)
	signal.Notify(sigCh, os.Interrupt, unix.SIGTERM)
	defer signal.Stop(sigCh)
	go func() {
		select {
		case <-ctx.Done():
		case <-sigCh:
			cancel()
		}
	}()

	exec := keyboard.ExecutorFunc(func(args []string) error {
		rootCmd.SetArgs(args)
		return rootCmd.ExecuteContext(ctx)
	})

	loop := keyboard.Loop{
		Executor:   exec,
		InfoWriter: c.ErrOrStderr(),
		KeyReader:  keyReaderWithCtrlArrows(exec),

		Bindings: keyboard.Bindings{
			Keys: map[githubkeyboard.Key]keyboard.Action{
				githubkeyboard.KeyArrowLeft:  keyboardAction(c, actions.Move, []string{"left"}),
				githubkeyboard.KeyArrowRight: keyboardAction(c, actions.Move, []string{"right"}),
				githubkeyboard.KeyArrowUp:    keyboardAction(c, actions.Move, []string{"up"}),
				githubkeyboard.KeyArrowDown:  keyboardAction(c, actions.Move, []string{"down"}),
				githubkeyboard.KeyEsc:        clearSelectionAction(c),
			},
			Runes: map[rune]keyboard.Action{
				'h': keyboardAction(c, actions.Move, []string{"left"}),
				'j': keyboardAction(c, actions.Move, []string{"down"}),
				'k': keyboardAction(c, actions.Move, []string{"up"}),
				'l': keyboardAction(c, actions.Move, []string{"right"}),
				'i': insertShortcut(c),
				'/': searchShortcut(c, false),
				'?': searchShortcut(c, true),
				'n': keyboardAction(c, actions.SearchRepeatForward, nil),
				'N': keyboardAction(c, actions.SearchRepeatBackward, nil),
				'v': visualRangeShortcut(c),
				'V': visualRowShortcut(c),
				's': keyboardAction(c, actions.Status, nil),
				'g': gotoShortcut(c),
				'c': columnHeaderShortcut(c),
				'r': rowHeaderShortcut(c),
				'y': yankShortcut(c),
				'Y': withRichTextWarning(c, hasRichTextRow, keyboardAction(c, actions.RowYank, nil)),
				'x': cutShortcut(c),
				'X': withRichTextWarning(c, hasRichTextRow, keyboardAction(c, actions.RowCut, nil)),
				'p': withRichTextWarning(c, hasRichTextSelectionOrCell, keyboardAction(c, actions.Paste, nil)),
				'P': withRichTextWarning(c, hasRichTextSelectionOrCell, keyboardAction(c, actions.Paste, []string{"--before"})),
				'd': deleteCellShortcut(c),
				'D': withRichTextWarning(c, hasRichTextRow, keyboardAction(c, actions.RowDelete, nil)),
				'O': keyboardAction(c, actions.RowInsertAbove, nil),
				'o': keyboardAction(c, actions.RowInsertBelow, nil),
			},
		},
	}

	loop.Terminal = ui.NewTerminal(int(os.Stdin.Fd()))

	for {
		if err := loop.Run(ctx); err != nil {
			if err == keyboard.ErrQuit {
				if appState.IsDirty() {
					fmt.Fprintln(c.OutOrStdout(), "No write since last change (add ! to override)")
					continue
				}
				fmt.Fprintln(c.OutOrStdout(), "exiting Ebenezer")
				return nil
			}
			return err
		}
		return nil
	}
}

var errPromptCanceled = errors.New("prompt cancelled")

var getKey = githubkeyboard.GetKey
var pollReadable = func(timeout time.Duration) (bool, error) {
	if timeout <= 0 {
		return true, nil
	}
	ms := int(timeout / time.Millisecond)
	if ms <= 0 {
		ms = 1
	}
	fds := []unix.PollFd{{
		Fd:     int32(os.Stdin.Fd()),
		Events: unix.POLLIN,
	}}
	n, err := unix.Poll(fds, ms)
	if err != nil {
		return false, err
	}
	if n == 0 {
		return false, nil
	}
	return fds[0].Revents&unix.POLLIN != 0, nil
}

type keyboardEvent struct {
	r rune
	k githubkeyboard.Key
}

// keyReaderWithCtrlArrows wraps github.com/eiannone/keyboard's GetKey and detects
// Ctrl+Arrow escape sequences (e.g. ESC [ 1 ; 5 D) that the library doesn't model
// as distinct keys. When detected, it executes the shared "move-span" command and
// returns an empty keypress so the normal dispatch layer ignores it.
func keyReaderWithCtrlArrows(exec keyboard.CommandExecutor) func() (rune, githubkeyboard.Key, error) {
	var pending []keyboardEvent
	const ctrlArrowWait = 25 * time.Millisecond

	type eventResult struct {
		ev  keyboardEvent
		err error
	}

	readNext := func(timeout time.Duration) (keyboardEvent, bool, error) {
		ready, err := pollReadable(timeout)
		if err != nil {
			return keyboardEvent{}, false, err
		}
		if !ready {
			return keyboardEvent{}, false, nil
		}
		r, k, err := getKey()
		if err != nil {
			return keyboardEvent{}, false, err
		}
		return keyboardEvent{r: r, k: k}, true, nil
	}

	return func() (rune, githubkeyboard.Key, error) {
		if len(pending) > 0 {
			ev := pending[0]
			pending = pending[1:]
			return ev.r, ev.k, nil
		}

		r, k, err := getKey()
		if err != nil {
			return 0, 0, err
		}
		ev := keyboardEvent{r: r, k: k}
		if k != githubkeyboard.KeyEsc {
			return r, k, nil
		}

		// Attempt to interpret Ctrl+Arrow escape sequences without blocking Esc.
		second, ok, err := readNext(ctrlArrowWait)
		if err != nil {
			return 0, 0, err
		}
		if !ok {
			return ev.r, ev.k, nil
		}

		// If this isn't an escape sequence start, buffer it and treat as plain Esc.
		if second.r != '[' {
			pending = append(pending, second)
			return ev.r, ev.k, nil
		}

		var seq []keyboardEvent
		for len(seq) < 8 {
			next, ok, err := readNext(ctrlArrowWait)
			if err != nil {
				return 0, 0, err
			}
			if !ok {
				break
			}
			seq = append(seq, next)
			// CSI sequences terminate with a final byte in the range 0x40-0x7E.
			if next.r >= '@' && next.r <= '~' {
				break
			}
		}

		// Build a minimal rune sequence we can match against.
		var b strings.Builder
		b.WriteRune('[')
		for _, ev := range seq {
			if ev.r == 0 {
				continue
			}
			b.WriteRune(ev.r)
		}
		if dir, ok := ctrlArrowDirection(b.String()); ok {
			_ = exec.ExecuteCommand([]string{"move-span", dir})
			return 0, 0, nil
		}

		// Unknown sequence: replay bytes after the initial Esc so they can be handled normally.
		pending = append(pending, second)
		pending = append(pending, seq...)
		return ev.r, ev.k, nil
	}
}

func ctrlArrowDirection(seq string) (string, bool) {
	if seq == "" {
		return "", false
	}
	if seq[0] == '[' {
		seq = seq[1:]
	}
	if len(seq) == 0 {
		return "", false
	}

	final := seq[len(seq)-1]
	switch final {
	case 'A', 'B', 'C', 'D':
	default:
		return "", false
	}

	params := seq[:len(seq)-1]
	if params == "" {
		return "", false
	}

	hasCtrl := false
	for _, part := range strings.Split(params, ";") {
		if part == "" {
			continue
		}
		val, err := strconv.Atoi(part)
		if err != nil {
			return "", false
		}
		if val == 5 {
			hasCtrl = true
		}
	}
	if !hasCtrl {
		return "", false
	}

	switch final {
	case 'A':
		return "up", true
	case 'B':
		return "down", true
	case 'C':
		return "right", true
	case 'D':
		return "left", true
	default:
		return "", false
	}
}

func gotoShortcut(c *cobra.Command) keyboard.Action {
	return func(ctx *keyboard.Context) error {
		char, key, err := getKey()
		if err != nil {
			return err
		}
		if dir, ok := moveSpanDirectionForKey(char, key); ok {
			return keyboardAction(c, actions.MoveSpan, []string{dir})(ctx)
		}
		addr, err := promptForAddressWithInitial(c, char, key)
		if err != nil {
			if errors.Is(err, errPromptCanceled) {
				return nil
			}
			return err
		}
		if addr == "" {
			return nil
		}
		_, err = executeAction(c, actions.Goto, []string{addr})
		return err
	}
}

func promptForAddress(c *cobra.Command) (string, error) {
	return promptForAddressWithInitial(c, 0, 0)
}

func promptForAddressWithInitial(c *cobra.Command, initialRune rune, initialKey githubkeyboard.Key) (string, error) {
	out := c.OutOrStdout()
	current := appState.Address()
	fmt.Fprintf(out, "\nGoto cell (ESC to cancel) [%s]: ", current)
	buffer := []rune{}
	if err := applyInitialAddressInput(out, initialRune, initialKey, &buffer); err != nil {
		if errors.Is(err, errPromptCanceled) {
			return "", err
		}
		return "", err
	}
	if initialKey == githubkeyboard.KeyEnter {
		if len(buffer) == 0 {
			return strings.TrimSpace(strings.ToUpper(current)), nil
		}
		return strings.TrimSpace(strings.ToUpper(string(buffer))), nil
	}
	for {
		char, key, err := getKey()
		if err != nil {
			return "", err
		}
		switch key {
		case githubkeyboard.KeyEsc:
			fmt.Fprintln(out)
			return "", errPromptCanceled
		case githubkeyboard.KeyEnter:
			fmt.Fprintln(out)
			if len(buffer) == 0 {
				return strings.TrimSpace(strings.ToUpper(current)), nil
			}
			return strings.TrimSpace(strings.ToUpper(string(buffer))), nil
		case githubkeyboard.KeyBackspace, githubkeyboard.KeyBackspace2:
			if len(buffer) > 0 {
				buffer = buffer[:len(buffer)-1]
				fmt.Fprint(out, "\b \b")
			}
		case githubkeyboard.KeyCtrlU, githubkeyboard.KeyCtrlW:
			handleEditingControl(out, key, &buffer)
			continue
		default:
			if unicode.IsLetter(char) {
				char = unicode.ToUpper(char)
				buffer = append(buffer, char)
				fmt.Fprint(out, string(char))
			} else if unicode.IsDigit(char) {
				buffer = append(buffer, char)
				fmt.Fprint(out, string(char))
			}
		}
	}
}

func applyInitialAddressInput(out io.Writer, char rune, key githubkeyboard.Key, buffer *[]rune) error {
	switch key {
	case githubkeyboard.KeyEsc:
		fmt.Fprintln(out)
		return errPromptCanceled
	case githubkeyboard.KeyEnter:
		fmt.Fprintln(out)
		return nil
	case githubkeyboard.KeyBackspace, githubkeyboard.KeyBackspace2:
		return nil
	case githubkeyboard.KeyCtrlU, githubkeyboard.KeyCtrlW:
		handleEditingControl(out, key, buffer)
		return nil
	default:
		if unicode.IsLetter(char) {
			char = unicode.ToUpper(char)
			*buffer = append(*buffer, char)
			fmt.Fprint(out, string(char))
		} else if unicode.IsDigit(char) {
			*buffer = append(*buffer, char)
			fmt.Fprint(out, string(char))
		}
	}
	return nil
}

func moveSpanDirectionForKey(char rune, key githubkeyboard.Key) (string, bool) {
	switch key {
	case githubkeyboard.KeyArrowLeft:
		return "left", true
	case githubkeyboard.KeyArrowRight:
		return "right", true
	case githubkeyboard.KeyArrowUp:
		return "up", true
	case githubkeyboard.KeyArrowDown:
		return "down", true
	}
	switch unicode.ToLower(char) {
	case 'h':
		return "left", true
	case 'l':
		return "right", true
	case 'k':
		return "up", true
	case 'j':
		return "down", true
	default:
		return "", false
	}
}

func columnHeaderShortcut(c *cobra.Command) keyboard.Action {
	return func(ctx *keyboard.Context) error {
		match, err := expectNextRune('t')
		if err != nil {
			return err
		}
		if !match {
			return nil
		}
		_, err = executeAction(c, actions.ColumnHeader, nil)
		return err
	}
}

func visualRangeShortcut(c *cobra.Command) keyboard.Action {
	return func(ctx *keyboard.Context) error {
		appState.ToggleSelection(app.SelectionRange)
		status.Print(c.OutOrStdout(), appState)
		return nil
	}
}

func visualRowShortcut(c *cobra.Command) keyboard.Action {
	return func(ctx *keyboard.Context) error {
		appState.ToggleSelection(app.SelectionRow)
		status.Print(c.OutOrStdout(), appState)
		return nil
	}
}

func clearSelectionAction(c *cobra.Command) keyboard.Action {
	return func(ctx *keyboard.Context) error {
		if appState.HasSelection() {
			appState.ClearSelection()
			status.Print(c.OutOrStdout(), appState)
		}
		return nil
	}
}

func keyboardAction(c *cobra.Command, action actions.Action, args []string) keyboard.Action {
	return func(ctx *keyboard.Context) error {
		ac := actions.NewContext(appState, c.OutOrStdout())
		ac.Logger = actions.NopLogger{}
		meta := action.Metadata()
		ac.Logger.Before(ac, meta, args)
		start := time.Now()
		res, err := action.Exec(ac, args)
		ac.Logger.After(ac, meta, args, res, err, time.Since(start))
		if err != nil {
			return err
		}
		if res.Message != "" {
			fmt.Fprint(c.OutOrStdout(), res.Message)
		}
		return err
	}
}

func rowHeaderShortcut(c *cobra.Command) keyboard.Action {
	return func(ctx *keyboard.Context) error {
		match, err := expectNextRune('t')
		if err != nil {
			return err
		}
		if !match {
			return nil
		}
		_, err = executeAction(c, actions.RowHeader, nil)
		return err
	}
}

func searchShortcut(c *cobra.Command, reverse bool) keyboard.Action {
	return func(ctx *keyboard.Context) error {
		initial := appState.LastSearchQuery()
		value, err := promptForSearch(c, initial, reverse)
		if err != nil {
			if errors.Is(err, errPromptCanceled) {
				return nil
			}
			return err
		}
		term := strings.TrimSpace(value)
		if term == "" {
			term = strings.TrimSpace(initial)
		}
		if term == "" {
			return nil
		}
		action := actions.SearchForward
		if reverse {
			action = actions.SearchBackward
		}
		return keyboardAction(c, action, []string{term})(ctx)
	}
}

func expectNextRune(target rune) (bool, error) {
	char, _, err := getKey()
	if err != nil {
		return false, err
	}
	return unicode.ToLower(char) == unicode.ToLower(target), nil
}

func nextRune() (rune, error) {
	char, _, err := getKey()
	return char, err
}

func insertShortcut(c *cobra.Command) keyboard.Action {
	return func(ctx *keyboard.Context) error {
		if hasRichTextCurrentCell() {
			ok, err := confirmRichTextFlatten(c)
			if err != nil || !ok {
				return err
			}
		}
		value, err := promptForText(c, appState.CurrentValue())
		if err != nil {
			if errors.Is(err, errPromptCanceled) {
				return nil
			}
			return err
		}
		return keyboardAction(c, actions.Edit, []string{value})(ctx)
	}
}

func deleteCellShortcut(c *cobra.Command) keyboard.Action {
	return func(ctx *keyboard.Context) error {
		char, err := nextRune()
		if err != nil {
			return err
		}
		switch unicode.ToLower(char) {
		case 'c':
			return withRichTextWarning(c, hasRichTextCurrentCell, keyboardAction(c, actions.Clear, nil))(ctx)
		case 'd':
			return withRichTextWarning(c, hasRichTextRow, keyboardAction(c, actions.RowDelete, nil))(ctx)
		case 'x':
			return withRichTextWarning(c, hasRichTextRow, keyboardAction(c, actions.RowCut, nil))(ctx)
		default:
			return nil
		}
	}
}

func yankShortcut(c *cobra.Command) keyboard.Action {
	return func(ctx *keyboard.Context) error {
		char, err := nextRune()
		if err != nil {
			return err
		}
		if unicode.ToLower(char) == 'y' {
			return withRichTextWarning(c, hasRichTextRow, keyboardAction(c, actions.RowYank, nil))(ctx)
		}
		return withRichTextWarning(c, hasRichTextCurrentCell, keyboardAction(c, actions.Yank, nil))(ctx)
	}
}

func cutShortcut(c *cobra.Command) keyboard.Action {
	return func(ctx *keyboard.Context) error {
		char, err := nextRune()
		if err != nil {
			return err
		}
		if unicode.ToLower(char) == 'x' {
			return withRichTextWarning(c, hasRichTextRow, keyboardAction(c, actions.RowCut, nil))(ctx)
		}
		return withRichTextWarning(c, hasRichTextCurrentCell, keyboardAction(c, actions.Cut, nil))(ctx)
	}
}

func withRichTextWarning(c *cobra.Command, hasRichText func() bool, action keyboard.Action) keyboard.Action {
	return func(ctx *keyboard.Context) error {
		if hasRichText != nil && hasRichText() {
			ok, err := confirmRichTextFlatten(c)
			if err != nil || !ok {
				return err
			}
		}
		return action(ctx)
	}
}

func confirmRichTextFlatten(c *cobra.Command) (bool, error) {
	out := c.OutOrStdout()
	fmt.Fprint(out, "\nWarning: this action will flatten rich text formatting. Proceed? (y/N): ")
	for {
		char, key, err := getKey()
		if err != nil {
			return false, err
		}
		switch key {
		case githubkeyboard.KeyEnter, githubkeyboard.KeyEsc:
			fmt.Fprintln(out)
			return false, nil
		}
		switch unicode.ToLower(char) {
		case 'y':
			fmt.Fprintln(out)
			return true, nil
		case 'n':
			fmt.Fprintln(out)
			return false, nil
		}
	}
}

func hasRichTextCurrentCell() bool {
	addr := strings.ToUpper(appState.Address())
	return richTextMapHasAddress(addr)
}

func hasRichTextSelectionOrCell() bool {
	if appState.HasSelection() {
		startRow, startCol, endRow, endCol, ok := appState.SelectionBounds()
		if ok {
			return richTextMapHasRange(startRow, startCol, endRow, endCol)
		}
	}
	return hasRichTextCurrentCell()
}

func hasRichTextRow() bool {
	if appState.HasSelection() && appState.Selection.Mode == app.SelectionRow {
		startRow, endRow, ok := appState.SelectionRowBounds()
		if ok {
			return richTextMapHasRowRange(startRow, endRow)
		}
	}
	return richTextMapHasRowRange(appState.Cursor.Row, appState.Cursor.Row)
}

func richTextMapHasAddress(addr string) bool {
	rt := richTextMapActive()
	if len(rt) == 0 || addr == "" {
		return false
	}
	_, ok := rt[addr]
	return ok
}

func richTextMapHasRange(startRow, startCol, endRow, endCol int) bool {
	rt := richTextMapActive()
	if len(rt) == 0 {
		return false
	}
	for addr := range rt {
		row, col, ok := parseCellAddress(addr)
		if !ok {
			continue
		}
		if row >= startRow && row <= endRow && col >= startCol && col <= endCol {
			return true
		}
	}
	return false
}

func richTextMapHasRowRange(startRow, endRow int) bool {
	rt := richTextMapActive()
	if len(rt) == 0 {
		return false
	}
	for addr := range rt {
		row, _, ok := parseCellAddress(addr)
		if !ok {
			continue
		}
		if row >= startRow && row <= endRow {
			return true
		}
	}
	return false
}

func richTextMapActive() map[string]int {
	if appState.Workbook == nil {
		return nil
	}
	if appState.Workbook.RichTextSheet != "" && !strings.EqualFold(appState.Workbook.RichTextSheet, appState.Workbook.Sheet) {
		return nil
	}
	return appState.Workbook.RichTextRuns
}

func parseCellAddress(addr string) (row, col int, ok bool) {
	addr = strings.TrimSpace(strings.ToUpper(addr))
	if addr == "" {
		return 0, 0, false
	}
	i := 0
	for i < len(addr) {
		ch := addr[i]
		if ch < 'A' || ch > 'Z' {
			break
		}
		col = col*26 + int(ch-'A'+1)
		i++
	}
	if i == 0 || i >= len(addr) {
		return 0, 0, false
	}
	rowVal, err := strconv.Atoi(addr[i:])
	if err != nil || rowVal < 1 {
		return 0, 0, false
	}
	return rowVal, col, true
}

func simpleCommand(name string, args ...string) keyboard.Action {
	return func(ctx *keyboard.Context) error {
		cmdArgs := append([]string{name}, args...)
		return ctx.Executor.ExecuteCommand(cmdArgs)
	}
}

func promptForText(c *cobra.Command, initial string) (string, error) {
	out := c.OutOrStdout()
	fmt.Fprintf(out, "\nEnter value (ESC to cancel) [%s]: ", initial)
	buffer := []rune(initial)
	fmt.Fprint(out, initial)
	for {
		char, key, err := githubkeyboard.GetKey()
		if err != nil {
			return "", err
		}
		switch key {
		case githubkeyboard.KeyEsc:
			fmt.Fprintln(out)
			return "", errPromptCanceled
		case githubkeyboard.KeyEnter:
			fmt.Fprintln(out)
			return string(buffer), nil
		case githubkeyboard.KeyBackspace, githubkeyboard.KeyBackspace2:
			if len(buffer) > 0 {
				buffer = buffer[:len(buffer)-1]
				fmt.Fprint(out, "\b \b")
			}
		case githubkeyboard.KeyCtrlU, githubkeyboard.KeyCtrlW:
			handleEditingControl(out, key, &buffer)
			continue
		default:
			if isPrintable(char) {
				buffer = append(buffer, char)
				fmt.Fprint(out, string(char))
			}
		}
	}
}

func isPrintable(r rune) bool {
	return r >= 32 && r != 127
}

func promptForSearch(c *cobra.Command, initial string, reverse bool) (string, error) {
	direction := "Forward"
	if reverse {
		direction = "Backward"
	}
	out := c.OutOrStdout()
	fmt.Fprintf(out, "\n%s search (ESC to cancel) [%s]: ", direction, initial)
	buffer := []rune(initial)
	fmt.Fprint(out, initial)
	for {
		char, key, err := githubkeyboard.GetKey()
		if err != nil {
			return "", err
		}
		switch key {
		case githubkeyboard.KeyEsc:
			fmt.Fprintln(out)
			return "", errPromptCanceled
		case githubkeyboard.KeyEnter:
			fmt.Fprintln(out)
			return string(buffer), nil
		case githubkeyboard.KeyBackspace, githubkeyboard.KeyBackspace2:
			if len(buffer) > 0 {
				buffer = buffer[:len(buffer)-1]
				fmt.Fprint(out, "\b \b")
			}
		case githubkeyboard.KeyCtrlU, githubkeyboard.KeyCtrlW:
			handleEditingControl(out, key, &buffer)
			continue
		default:
			if isPrintable(char) {
				buffer = append(buffer, char)
				fmt.Fprint(out, string(char))
			}
		}
	}
}

func handleEditingControl(out io.Writer, key githubkeyboard.Key, buffer *[]rune) {
	switch key {
	case githubkeyboard.KeyCtrlU:
		removed := len(*buffer)
		*buffer = (*buffer)[:0]
		eraseChars(out, removed)
	case githubkeyboard.KeyCtrlW:
		removed := deleteTrailingWord(buffer)
		eraseChars(out, removed)
	}
}

func deleteTrailingWord(buffer *[]rune) int {
	b := *buffer
	if len(b) == 0 {
		return 0
	}
	i := len(b)
	for i > 0 && unicode.IsSpace(b[i-1]) {
		i--
	}
	for i > 0 && !unicode.IsSpace(b[i-1]) {
		i--
	}
	removed := len(b) - i
	*buffer = b[:i]
	return removed
}

func eraseChars(out io.Writer, count int) {
	for i := 0; i < count; i++ {
		fmt.Fprint(out, "\b \b")
	}
}
