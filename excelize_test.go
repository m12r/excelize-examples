package excelize_examples

import (
	"testing"

	"github.com/xuri/excelize/v2"

	"github.com/m12r/excelize-examples/internal/excelizetest"
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
	excelizetest.Dump(t, x, fileOpts)
}

func TestSharedFormulaWithDuplicateRow(t *testing.T) {
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
	if err := x.DuplicateRow(sheetName, 10); err != nil {
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
	if err := x.SetCellStyle(sheetName, "A1", "C11", styleID); err != nil {
		t.Fatal(err)
	}

	for row := 1; row <= 11; row++ {
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
	excelizetest.Dump(t, x, fileOpts)
}
