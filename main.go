package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"io"
	"log"
	"os"
	"runtime"
)

var (
	csvPathFlag   = flag.String("f", "", "Path to the CSV input file")
	delimiterFlag = flag.String("d", ",", "Delimiter for fields in the CSV input.")
	libNameFlag   = flag.String("l", "excelize", "lib name")
)

const (
	excelizeLibName   = "excelize"
	excelizeV2LibName = "excelizeV2"
	tealegLibName     = "tealeg"
)

func usage() {
	fmt.Printf(`%s: -f=<CSV Input File> -d=<Delimiter> -l=<Lib name>
  Lib names:
    - %s
    - %s
    - %s
`, os.Args[0], excelizeLibName, excelizeV2LibName, tealegLibName)
}

func main() {
	PrintMemUsage("start")

	flag.Parse()
	if len(os.Args) < 2 {
		usage()
		return
	}

	err := csv2xlsx(*csvPathFlag, *delimiterFlag, *libNameFlag)
	if err != nil {
		log.Fatal("Error: ", err)
	}

	PrintMemUsage("before runtime.GC()")
	runtime.GC()

	PrintMemUsage("end")
}

type converter interface {
	AddRow([]string) error
	Save(path string) error
}

func csv2xlsx(csvPath, delimiter, libName string) error {
	csvFile, err := os.Open(csvPath)
	if err != nil {
		return fmt.Errorf("failed to open %q: %w", csvPath, err)
	}
	defer csvFile.Close()

	reader := csv.NewReader(csvFile)
	if len(delimiter) > 0 {
		reader.Comma = rune((delimiter)[0])
	}

	var conv converter
	switch *libNameFlag {
	case excelizeLibName:
		conv = NewExcelizeConvertor()
	case excelizeV2LibName:
		conv = NewExcelizeV2Convertor()
	case tealegLibName:
		conv = NewTealegConvertor()
	default:
		return fmt.Errorf("unknown lib name %q", *libNameFlag)
	}

	xlsxPath := csvPath + "." + libName + ".xlsx"
	err = convert(reader, conv, xlsxPath)
	if err != nil {
		return fmt.Errorf("failed to convert %q: %w", csvPath, err)
	}

	return nil
}

func convert(reader *csv.Reader, conv converter, xlsxPath string) error {
	PrintMemUsage("before adding rows")
	for {
		record, err := reader.Read()
		if err != nil {
			if err != io.EOF {
				return fmt.Errorf("failed to read record: %w", err)
			}
			break
		}

		err = conv.AddRow(record)
		if err != nil {
			return fmt.Errorf("failed to add row: %w", err)
		}
	}

	PrintMemUsage("before save")
	err := conv.Save(xlsxPath)
	if err != nil {
		return fmt.Errorf("failed to save %q: %w", xlsxPath, err)
	}

	return nil
}

// PrintMemUsage outputs the current, total and OS memory being used. As well as the number
// of garbage collection cycles completed.
func PrintMemUsage(msg string) {
	var m runtime.MemStats
	runtime.ReadMemStats(&m)
	// For info on each, see: https://golang.org/pkg/runtime/#MemStats
	fmt.Printf("Alloc = %v MB", bToMb(m.Alloc))
	fmt.Printf("\tTotalAlloc = %v MB", bToMb(m.TotalAlloc))
	fmt.Printf("\tSys = %v MB", bToMb(m.Sys))
	fmt.Printf("\tNumGC = %v", m.NumGC)
	fmt.Printf("\t%v\n", msg)
}

func bToMb(b uint64) uint64 {
	return b / 1024 / 1024
}
