package main

import (
	"fmt"

	"github.com/plandem/xlsx"
)

func NewPlandemConvertor() *PlandemConvertor {
	return &PlandemConvertor{
		xlsxFile:  xlsx.New(),
		sheetName: "Sheet1",
		rowIndex:  0,
	}
}

type PlandemConvertor struct {
	xlsxFile  *xlsx.Spreadsheet
	sheetName string
	sheet     xlsx.Sheet
	rowIndex  int
}

func (e *PlandemConvertor) AddRow(row []string) error {
	if e.sheet == nil {
		e.sheet = e.xlsxFile.AddSheet(e.sheetName)
	}

	for colIndex, field := range row {
		cell := e.sheet.Cell(colIndex, e.rowIndex)
		cell.SetValue(field)
	}

	e.rowIndex++
	return nil
}

func (e *PlandemConvertor) Save(path string) error {
	err := e.xlsxFile.SaveAs(path)
	if err != nil {
		return fmt.Errorf("failed to save %q: %w", path, err)
	}
	return nil
}

func (e *PlandemConvertor) Close() error {
	err := e.xlsxFile.Close()
	if err != nil {
		return fmt.Errorf("failed to close file: %w", err)
	}
	return nil
}
