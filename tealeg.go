package main

import (
	"fmt"

	"github.com/tealeg/xlsx/v3"
)

func NewTealegConvertor() *TealegConvertor {
	return &TealegConvertor{
		xlsxFile:  xlsx.NewFile(),
		sheetName: "Sheet1",
	}
}

type TealegConvertor struct {
	xlsxFile  *xlsx.File
	sheetName string
	sheet     *xlsx.Sheet
}

func (e *TealegConvertor) AddRow(row []string) error {
	if e.sheet == nil {
		sheet, err := e.xlsxFile.AddSheet(e.sheetName)
		if err != nil {
			return fmt.Errorf("failed to AddSheet %q: %w", e.sheetName, err)
		}
		e.sheet = sheet
	}

	sheetRow := e.sheet.AddRow()
	for _, field := range row {
		cell := sheetRow.AddCell()
		cell.Value = field
	}

	return nil
}

func (e *TealegConvertor) Save(path string) error {
	err := e.xlsxFile.Save(path)
	if err != nil {
		return fmt.Errorf("failed to save %q: %w", path, err)
	}
	return nil
}
