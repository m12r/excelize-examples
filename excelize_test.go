package excelize_examples

import (
	"fmt"
	"os"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestSharedFormula(t *testing.T) {
	fileOpts := excelize.Options{
		RawCellValue: false,
	}
	x := excelize.NewFile(fileOpts)
	sheetName := x.GetSheetName(0)
	for row := 1; row <= 10; row++ {
		cell := cellAddr("A", row)
		if err := x.SetCellFloat(sheetName, cell, 1.50, 2, 64); err != nil {
			t.Fatal(err)
		}
		cell = cellAddr("B", row)
		if err := x.SetCellFloat(sheetName, cell, 2.49, 2, 64); err != nil {
			t.Fatal(err)
		}
	}
	formulaType, formulaRef := excelize.STCellFormulaTypeShared, "C1:C10"
	if err := x.SetCellFormula(sheetName, "C1", "A1+B1", excelize.FormulaOpts{Type: &formulaType, Ref: &formulaRef}); err != nil {
		t.Fatal(err)
	}
	decimalPlaces, customNumFmt := 2, "0.00"
	styleID, err := x.NewStyle(&excelize.Style{
		DecimalPlaces: &decimalPlaces,
		CustomNumFmt:  &customNumFmt,
	})
	if err != nil {
		t.Fatal(err)
	}
	if err := x.SetCellStyle(sheetName, "A1", "C10", styleID); err != nil {
		t.Fatal(err)
	}
	if err := x.ProtectSheet(
		sheetName,
		&excelize.SheetProtectionOptions{
			SelectLockedCells: true,
		},
	); err != nil {
		t.Fatal(err)
	}

	for row := 1; row <= 10; row++ {
		cell := cellAddr("C", row)
		formula, err := x.GetCellFormula(sheetName, cell)
		if err != nil {
			t.Fatal(err)
		}
		t.Log(formula)
	}

	if err := x.UpdateLinkedValue(); err != nil {
		t.Fatal(err)
	}
	f, err := os.OpenFile("test-shared-formula.xlsx", os.O_CREATE|os.O_TRUNC|os.O_WRONLY, 0o644)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	if _, err := x.WriteTo(f, fileOpts); err != nil {
		t.Fatal(err)
	}
}

func cellAddr(column string, row int) string {
	return fmt.Sprintf("%s%d", column, row)
}
