package workbook

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"testing"

	"github.com/google/go-cmp/cmp"
)

const goldenDir = "testdata"

func TestSaveCSVGolden(t *testing.T) {
	wb := SampleWorkbook()
	tmp := filepath.Join(t.TempDir(), "out.csv")
	if err := wb.Save(tmp); err != nil {
		t.Fatalf("save csv for golden test failed: %v", err)
	}
	data, err := os.ReadFile(tmp)
	if err != nil {
		t.Fatalf("read csv output: %v", err)
	}
	assertGoldenBytes(t, "save.csv", data)
}

func TestSaveXLSXGolden(t *testing.T) {
	wb := styledWorkbookForGolden()
	tmp := filepath.Join(t.TempDir(), "styled.xlsx")
	if err := wb.Save(tmp); err != nil {
		t.Fatalf("save xlsx for golden test failed: %v", err)
	}
	compareXLSXGolden(t, tmp, "styled.xlsx")
}

func styledWorkbookForGolden() *Workbook {
	wb := SampleWorkbook()
	wb.ColumnWidths = map[int]float64{
		1: 18.5,
		2: 16,
		3: 12,
	}
	wb.Styles["A1"] = CellStyle{
		FillColor: "FFD965",
		Bold:      true,
		Borders: map[string]BorderStyle{
			"bottom": {Style: "thin", Color: "000000"},
		},
	}
	wb.Styles["B2"] = CellStyle{
		FontColor:       "FF0000",
		Italic:          true,
		HorizontalAlign: "center",
	}
	wb.Styles["C5"] = CellStyle{
		NumberFormat: "builtin:4",
		Underline:    true,
	}
	wb.ActiveCell = "C3"
	return wb
}

func assertGoldenBytes(t *testing.T, name string, actual []byte) {
	t.Helper()
	path := filepath.Join(goldenDir, name)
	if updateGoldenFiles() {
		if err := os.WriteFile(path, actual, 0o644); err != nil {
			t.Fatalf("overwrite golden %s: %v", name, err)
		}
		t.Logf("updated golden %s", name)
		return
	}
	expected, err := os.ReadFile(path)
	if err != nil {
		t.Fatalf("read golden %s: %v", name, err)
	}
	if !bytes.Equal(expected, actual) {
		delta := cmp.Diff(string(expected), string(actual))
		t.Fatalf("csv output differs from golden %s:\n%s", name, delta)
	}
}

func compareXLSXGolden(t *testing.T, gotPath, name string) {
	t.Helper()
	path := filepath.Join(goldenDir, name)
	if updateGoldenFiles() {
		if err := copyFile(path, gotPath); err != nil {
			t.Fatalf("overwrite golden %s: %v", name, err)
		}
		t.Logf("updated golden %s", name)
		return
	}
	diff, err := zippedXLSXDiff(gotPath, path)
	if err != nil {
		t.Fatalf("compare xlsx golden %s: %v", name, err)
	}
	if diff != "" {
		t.Fatalf("xlsx output differs from golden %s:\n%s", name, diff)
	}
}

func copyFile(dst, src string) error {
	data, err := os.ReadFile(src)
	if err != nil {
		return err
	}
	return os.WriteFile(dst, data, 0o644)
}

func zippedXLSXDiff(gotPath, goldenPath string) (string, error) {
	got, err := xlsxHashes(gotPath)
	if err != nil {
		return "", fmt.Errorf("hash actual xlsx: %w", err)
	}
	want, err := xlsxHashes(goldenPath)
	if err != nil {
		return "", fmt.Errorf("hash golden xlsx: %w", err)
	}
	return diffHashMaps(got, want), nil
}

func xlsxHashes(path string) (map[string]string, error) {
	reader, err := zip.OpenReader(path)
	if err != nil {
		return nil, fmt.Errorf("open xlsx %s: %w", path, err)
	}
	defer reader.Close()
	hashes := map[string]string{}
	for _, file := range reader.File {
		if ignoreXLSXEntry(file.Name) {
			continue
		}
		rc, err := file.Open()
		if err != nil {
			return nil, err
		}
		data, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			return nil, err
		}
		sum := sha256.Sum256(data)
		hashes[file.Name] = fmt.Sprintf("%x", sum)
	}
	return hashes, nil
}

func diffHashMaps(got, want map[string]string) string {
	keys := map[string]struct{}{}
	for k := range got {
		keys[k] = struct{}{}
	}
	for k := range want {
		keys[k] = struct{}{}
	}
	if len(keys) == 0 {
		return ""
	}
	var sorted []string
	for k := range keys {
		sorted = append(sorted, k)
	}
	sort.Strings(sorted)
	var diffs []string
	for _, key := range sorted {
		actual, aok := got[key]
		expected, eok := want[key]
		switch {
		case !aok:
			diffs = append(diffs, fmt.Sprintf("missing entry in actual: %s", key))
		case !eok:
			diffs = append(diffs, fmt.Sprintf("extra entry in actual: %s", key))
		case actual != expected:
			diffs = append(diffs, fmt.Sprintf("hash mismatch %s: actual=%s expected=%s", key, actual[:8], expected[:8]))
		}
	}
	return strings.Join(diffs, "\n")
}

func ignoreXLSXEntry(name string) bool {
	lower := strings.ToLower(name)
	switch lower {
	case "docprops/core.xml", "docprops/app.xml":
		return true
	}
	return false
}

func updateGoldenFiles() bool {
	return os.Getenv("UPDATE_GOLDEN") == "1"
}
