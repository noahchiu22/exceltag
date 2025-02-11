package exceltag

import (
	"fmt"
	"reflect"
	"time"
	"unicode"
	"unicode/utf8"

	"github.com/xuri/excelize/v2"
)

// Using any excel tag as title in the struct to create excel
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
func CreateExcel[S comparable](data []S, path, filename string) error {
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

			if !field.CanSet() {
				err := fmt.Errorf("filed name have to be capitalize")
				return err
			}

			value := field.Interface()
			switch temp := value.(type) {
			case int64:
				value = temp
			case float64:
				value = temp
			case time.Time:
				value = temp.Format("2006-01-02 15:04:05")
			default:
				value = temp
			}
			f.SetCellValue("Sheet1", cell, value)
		}
	}

	err := AutofitColumn(f)
	if err != nil {
		return err
	}

	if path == "" {
		path = "./"
	}
	// Save spreadsheet by the given path.
	if err := f.SaveAs(path + filename); err != nil {
		return err
	}

	return nil
}

// Autofit all columns according to their text content
func AutofitColumn(file *excelize.File) error {
	cols, _ := file.GetCols("Sheet1")
	for i, col := range cols {
		largestWidth := 0
		for j, rowCell := range col {
			fixedWidth := 0
			for _, r := range rowCell {
				if unicode.Is(unicode.Han, r) {
					fixedWidth++
				}
			}
			cellWidth := utf8.RuneCountInString(rowCell) + 4 + fixedWidth // for margin
			if cellWidth > largestWidth {
				largestWidth = cellWidth
			}
			file.SetRowHeight("Sheet1", j+1, 25)
		}
		name, err := excelize.ColumnNumberToName(i + 1)
		if err != nil {
			return err
		}
		file.SetColWidth("Sheet1", name, name, float64(largestWidth))
	}

	return nil
}
