# exceltag
Generate a simple excel file with tags in your struct

```go
// Using any excel tag as title in the struct to create excel
//
//	type example struct {
//		fieldName  any  `excel:"your title"`
//	}
//
// Excel will be:
//
//	---------------
//	|  title  | ...
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
func CreateExcel[S comparable](data []S, path, filename string) (err error) {
	f := excelize.NewFile()

	defer func() {
		if err := f.Close(); err != nil {
			return
		}
	}()

	headers := []string{}
	for row, item := range data {
		t := reflect.TypeOf(item)
		// 第一筆資料依照struct的tag excel製作headers
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
			// 沒有excelTag就跳過
			if excelTag == "" {
				continue
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

	AutofitColumn(f)

	fmt.Println(path + filename)
	// Save spreadsheet by the given path.
	if err = f.SaveAs(path + filename); err != nil {
		return
	}

	return
}
```
