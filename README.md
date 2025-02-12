# exceltag
## Generate a simple excel file with tags in your struct

This package will generate a simple excel by your array of any struct with adding tag `excel:"your header"`

Futhermore, it will ignore field without excel tag if you don't want to export to excel

For example:
```go

type studentInfo struct {
	Num   int64  `excel:"student number"`
	Name  string `excel:"student name"`
	UpdatedAt time.Time `excel:"updated time"`
	// other field without excel tag if you don't want to export to excel
	other string
}

func main() {
	students := []studentInfo{
		{Num: 1, Name: "John", UpdatedAt: time.Now(), other: "other"},
		{Num: 2, Name: "Jane", UpdatedAt: time.Now(), other: "other"},
		{Num: 3, Name: "Jim", UpdatedAt: time.Now(), other: "other"},
		{Num: 4, Name: "Jill", UpdatedAt: time.Now(), other: "other"},
		{Num: 5, Name: "Jack", UpdatedAt: time.Now(), other: "other"},
		{Num: 6, Name: "Jill", UpdatedAt: time.Now(), other: "other"},
	}
	// generate excel file
	f, err := exceltag.CreateExcel(students)
	if err != nil {
		fmt.Println("err", err)
		return
	}

	// you can add some change or style to your excel file...

	// Save spreadsheet by the given path.
	if err := f.SaveAs("studentInfo.xlsx"); err != nil {
		fmt.Println("err", err)
		return
	}
}
```

and the simple spreadsheet of student data will be generated as follows:

![image](https://github.com/user-attachments/assets/b9f46054-60a1-4a35-8f5d-c783a733e58d)


## Customizing the Excel File with excelize
Using the excelize's functions to customize the excel file, like changing the sheet name, column width, and row color.

You can change the sheet name by using the excelize's `SetSheetName` function:

```go
err := f.SetSheetName("Sheet1", "Student Info")
if err != nil {
	fmt.Println("err", err)
	return
}
```

also, you can change the column width by using the excelize's `SetColWidth` function:

```go
err = f.SetColWidth("Sheet1", "A", "B", 20)
if err != nil {
	fmt.Println("err", err)
	return
}
```

and row color by using the excelize's `SetRowStyle` function:

```go
styleID, err := f.NewStyle(&excelize.Style{
	Fill: excelize.Fill{
		Type:    "pattern",
		Color:   []string{"#C0DCF4"},
		Pattern: 1,
	},
})
if err != nil {	
	fmt.Println("err", err)
	return
}

err = f.SetRowStyle("Sheet1", 1, 1, styleID)
if err != nil {
	fmt.Println("err", err)
	return
}
```

## Autofit column width

The function `AutofitColumn` will automatically fit the column width by the cell content, you can use the `AutofitColumn` function to do this.

```go
err := exceltag.AutofitColumn(f, "Sheet1")
if err != nil {
	fmt.Println("err", err)
	return
}
```

# Resource
[qax-os/excelize
](https://github.com/qax-os/excelize)
