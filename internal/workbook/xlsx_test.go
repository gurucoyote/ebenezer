package workbook

import (
	"archive/zip"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestFromXLSXIgnoresIncompleteTheme(t *testing.T) {
	path := buildBrokenThemeWorkbook(t)

	if _, _, _, err := FromXLSX(path, ""); err != nil {
		t.Fatalf("FromXLSX returned error for broken theme file: %v", err)
	}
}

func buildBrokenThemeWorkbook(t *testing.T) string {
	t.Helper()

	// Create a minimal workbook so we can corrupt its theme and styles.
	base := filepath.Join(t.TempDir(), "base.xlsx")
	f := excelize.NewFile()
	if err := f.SetCellValue("Sheet1", "A1", "themed"); err != nil {
		t.Fatalf("set cell: %v", err)
	}
	if err := f.SaveAs(base); err != nil {
		t.Fatalf("save base: %v", err)
	}

	broken := filepath.Join(t.TempDir(), "broken-theme.xlsx")
	rewriteWithBrokenTheme(t, base, broken)
	return broken
}

func rewriteWithBrokenTheme(t *testing.T, src, dst string) {
	t.Helper()

	reader, err := zip.OpenReader(src)
	if err != nil {
		t.Fatalf("open base zip: %v", err)
	}
	defer reader.Close()

	outFile, err := os.Create(dst)
	if err != nil {
		t.Fatalf("create broken zip: %v", err)
	}
	defer outFile.Close()

	writer := zip.NewWriter(outFile)
	defer writer.Close()

	for _, file := range reader.File {
		header := &zip.FileHeader{
			Name:   file.Name,
			Method: zip.Deflate,
		}
		// Preserve directory entries as-is.
		if strings.HasSuffix(file.Name, "/") {
			header.Method = zip.Store
		}

		w, err := writer.CreateHeader(header)
		if err != nil {
			t.Fatalf("create header for %s: %v", file.Name, err)
		}

		switch file.Name {
		case "xl/theme/theme1.xml":
			// Provide a theme node with no theme elements, which causes
			// excelize to panic when resolving theme colors unless sanitized.
			_, _ = io.WriteString(w, `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></a:theme>`)
		case "xl/styles.xml":
			data, err := readZipFile(file)
			if err != nil {
				t.Fatalf("read styles.xml: %v", err)
			}
			data = strings.Replace(data,
				`<patternFill patternType="gray125"></patternFill>`,
				`<patternFill patternType="solid"><fgColor theme="1"/></patternFill>`, 1)
			data = strings.Replace(data,
				`cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0">`,
				`cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="1" borderId="0" xfId="0">`, 1)
			if _, err := io.WriteString(w, data); err != nil {
				t.Fatalf("write patched styles.xml: %v", err)
			}
		default:
			if err := copyZipEntry(w, file); err != nil {
				t.Fatalf("copy %s: %v", file.Name, err)
			}
		}
	}
}

func readZipFile(f *zip.File) (string, error) {
	rc, err := f.Open()
	if err != nil {
		return "", err
	}
	defer rc.Close()
	data, err := io.ReadAll(rc)
	if err != nil {
		return "", err
	}
	return string(data), nil
}

func copyZipEntry(dst io.Writer, src *zip.File) error {
	rc, err := src.Open()
	if err != nil {
		return err
	}
	defer rc.Close()
	_, err = io.Copy(dst, rc)
	return err
}
