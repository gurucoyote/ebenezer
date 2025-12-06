# Ebenezer CLI Demo

This walkthrough exercises the current Go skeleton so you can see the keyboard loop drive simple workbook navigation.

## Build
```
go build ./cmd/ebenezer
```

## Run with Sample Workbook
```
./ebenezer keyboard
```

### Normal Mode Bindings (current scope)
- Arrow keys: move the cursor (updates stdout with coordinates/value).
- `s`: print the current cell status.
- `:`: enter command mode (type commands such as `status`, `open test.csv`, `sample`).
- `q`: exit keyboard mode.

### Command Examples
```
:status          # show active cell/value
:open js/test.csv  # load an existing CSV in the repo
:sample          # reload built-in sample workbook
```

These commands exercise the workbook loader (`internal/workbook`), shared app state (`internal/app`), and the keyboard loop (`internal/ui/keyboard`).
