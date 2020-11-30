package main

import (
	"fmt"
	"os"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

func NewExcelizeV2Convertor() *ExcelizeV2Convertor {
	return &ExcelizeV2Convertor{
		xlsxFile:  excelize.NewFile(),
		sheetName: "Sheet1",
		rowIndex:  1,
	}
}

type ExcelizeV2Convertor struct {
	xlsxFile  *excelize.File
	sheetName string
	rowIndex  int
}

func (e *ExcelizeV2Convertor) AddRow(row []string) error {
	cell := fmt.Sprintf("A%d", e.rowIndex)

	err := e.xlsxFile.SetSheetRow(e.sheetName, cell, &row) // for v2.3.1
	if err != nil {
		return fmt.Errorf("failed to SetSheetRow: %w", err)
	}

	e.rowIndex++
	return nil
}

func (e ExcelizeV2Convertor) Save(path string) error {
	file, err := os.Create(path)
	if err != nil {
		return fmt.Errorf("failed to create %q: %w", path, err)
	}
	defer file.Close()

	err = e.xlsxFile.Write(file)
	if err != nil {
		return fmt.Errorf("failed to write %q: %w", path, err)
	}
	return nil
}
