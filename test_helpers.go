package excelize_examples

import (
	"io"
	"os"
	"path"
	"strings"
	"testing"

	"github.com/iancoleman/strcase"
	"github.com/xuri/excelize/v2"
)

func dumpExcelizeFile(t *testing.T, x *excelize.File, opts ...excelize.Options) {
	t.Helper()

	dumpPath := "_dump"
	ensurePath(t, dumpPath)

	basename, _ := strings.CutPrefix(strcase.ToKebab(t.Name()), "test-")
	fileExtensions := []string{".xlsx", ".zip"}
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

func ensurePath(t *testing.T, pathname string) {
	t.Helper()

	var err error
	_, err = os.Stat(pathname)
	if os.IsNotExist(err) {
		err = os.MkdirAll(pathname, 0o755)
	}
	if err != nil {
		t.Fatal(err)
	}
}
