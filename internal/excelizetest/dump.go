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

func Dump(t *testing.T, x *excelize.File, opts ...excelize.Options) {
	t.Helper()

	dumpPath := "_dump"
	basename, _ := strings.CutPrefix(strcase.ToKebab(t.Name()), "test-")
	fileExtensions := []string{".xlsx", ".zip"}

	if err := ensurePath(dumpPath); err != nil {
		t.Fatal(err)
	}

	writers := make([]io.Writer, 0, len(fileExtensions))

	for _, extension := range fileExtensions {
		filename := path.Join(dumpPath, basename+extension)

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
	if _, err := x.WriteTo(w, opts...); err != nil {
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
