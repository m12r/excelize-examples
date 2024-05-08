package excelize_examples

import "fmt"

func cellAddr(column string, row int) string {
	return fmt.Sprintf("%s%d", column, row)
}
