package exceltag

import (
	"fmt"
	"reflect"
	"unicode"
	"unicode/utf8"

	"github.com/xuri/excelize/v2"
)

// Using any excel tag as title in the struct to create a new excel file
//
//	type example struct {
//		fieldName  any  `excel:"your header"`
//	}
//
// Excel will be:
//
//	---------------
//	|  header  | ...
//	|---------+----
//	| data[i] | ...
//	|---------+----
//	|data[i+1]| ...
//	|---------+----
//	|data[i+2]| ...
//	|---------+----
//	|    :    | ...
//
// And it will automatically fit column width by the cell content
func CreateExcel[S comparable](data []S) (*excelize.File, error) {
	f := excelize.NewFile()

	defer func() error {
		if err := f.Close(); err != nil {
			return err
		}
		return nil
	}()

	headers := []string{}
	for row, item := range data {
		t := reflect.TypeOf(item)
		// Create headers from excel tags in struct for first row
		if row == 0 {
			for i := 0; i < t.NumField(); i++ {
				excelTag := t.Field(i).Tag.Get("excel")
				if excelTag != "" {
					headers = append(headers, excelTag)
				}
			}

			for col, header := range headers {
				cell, _ := excelize.CoordinatesToCellName(col+1, 1)
				f.SetCellValue("Sheet1", cell, header)
			}
		}
		v := reflect.ValueOf(item)
		for col := 0; col < len(headers); col++ {
			cell, _ := excelize.CoordinatesToCellName(col+1, row+2)
			field := v.Field(col)
			excelTag := t.Field(col).Tag.Get("excel")
			// Skip if no excel tag
			if excelTag == "" {
				continue
			}

			if !field.CanInterface() {
				return f, fmt.Errorf("filed name have to be capitalize")
			}

			value := field.Interface()
			err := f.SetCellValue("Sheet1", cell, value)
			if err != nil {
				return f, err
			}
		}
	}

	err := AutofitColumn(f, "Sheet1")
	if err != nil {
		return f, err
	}

	return f, nil
}

// Autofit all columns according to their text content of the given sheet
func AutofitColumn(file *excelize.File, sheetName string) error {
	cols, _ := file.GetCols(sheetName)
	for i, col := range cols {
		largestWidth := 0
		for _, rowCell := range col {
			fixedWidth := 0
			// check if the cell is chinese or japanese
			for _, r := range rowCell {
				if unicode.Is(unicode.Han, r) {
					fixedWidth++
				}
			}
			cellWidth := utf8.RuneCountInString(rowCell) + fixedWidth + 4 // for margin
			if cellWidth > largestWidth {
				largestWidth = cellWidth
			}
		}
		name, err := excelize.ColumnNumberToName(i + 1)
		if err != nil {
			return err
		}
		file.SetColWidth(sheetName, name, name, float64(largestWidth))
	}

	return nil
}
