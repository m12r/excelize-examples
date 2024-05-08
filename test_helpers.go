package excelize_examples

import (
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

	basename, _ := strings.CutPrefix(t.Name(), "Test")
	filename := path.Join(dumpPath, strcase.ToKebab(basename)+".dump.xlsx")

	f, err := os.OpenFile(filename, os.O_CREATE|os.O_TRUNC|os.O_WRONLY, 0o644)
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() {
		_ = f.Close()
	})

	if _, err := x.WriteTo(f, opts...); err != nil {
		t.Fatal(err)
	}
}

func ensurePath(t *testing.T, pathname string) {
	t.Helper()

	var err error
	_, err = os.Stat(pathname)
	if os.IsNotExist(err) {
		err = os.MkdirAll(pathname, os.ModePerm)
	}
	if err != nil {
		t.Fatal(err)
	}
}
