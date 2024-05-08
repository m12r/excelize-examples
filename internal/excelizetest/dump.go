package excelizetest

import (
	"io"
	"os"
	"path"
	"strings"
	"testing"

	"github.com/iancoleman/strcase"
	"github.com/xuri/excelize/v2"
)

type DumpParams struct {
	File    *excelize.File
	Options excelize.Options
}

type DumpOptions struct {
	Path       string
	Extensions []string
}

func DefaultDumpOptions() DumpOptions {
	extensions := make([]string, len(defaultDumpOptions.Extensions))
	copy(extensions, defaultDumpOptions.Extensions)

	return DumpOptions{
		Path:       defaultDumpOptions.Path,
		Extensions: extensions,
	}
}

var defaultDumpOptions = DumpOptions{
	Path:       "_dump",
	Extensions: []string{".xlsx", ".zip"},
}

func Dump(t *testing.T, params *DumpParams, opts DumpOptions) {
	t.Helper()

	if len(opts.Extensions) == 0 {
		t.Log("no extensions")
		return
	}

	basename, _ := strings.CutPrefix(strcase.ToKebab(t.Name()), "test-")

	if err := ensurePath(opts.Path); err != nil {
		t.Fatal(err)
	}

	writers := make([]io.Writer, 0, len(opts.Extensions))

	for _, extension := range opts.Extensions {
		filename := path.Join(opts.Path, basename+extension)

		f, err := os.OpenFile(filename, os.O_CREATE|os.O_TRUNC|os.O_WRONLY, 0o644)
		if err != nil {
			t.Fatal(err)
		}

		writers = append(writers, f)

		t.Cleanup(func() {
			_ = f.Close()
		})
	}

	w := io.MultiWriter(writers...)
	if _, err := params.File.WriteTo(w, params.Options); err != nil {
		t.Fatal(err)
	}
}

func ensurePath(pathname string) (err error) {
	_, err = os.Stat(pathname)
	if os.IsNotExist(err) {
		err = os.MkdirAll(pathname, 0o755)
	}
	return err
}
