package excelize_examples

import (
	"io"
	"os"
	"testing"
)

func openFileAsWriter(t *testing.T, filename string) io.Writer {
	t.Helper()

	f, err := os.OpenFile(filename, os.O_CREATE|os.O_TRUNC|os.O_WRONLY, 0o644)
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() {
		_ = f.Close()
	})
	return f
}
