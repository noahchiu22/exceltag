# exceltag
## Generate a simple excel file with tags in your struct

This package will generate a simple excel by your array of any struct with adding tag `excel:"your header"`

For example:
```go
type example struct {
	FieldName  any  `excel:"header"`
	Other any // you can put other field without excel tag if you don't want to export to excel
}

func main() {
	examples := []example{
		{FieldName: "first item"},
		{FieldName: "second item"},
		{FieldName: "third item"},
	}
	err := exceltag.CreateExcel(examples, "./", "example.xlsx")
	if err != nil {
		fmt.Println("err", err)
	}
}
```

will generate an excel named as example.xlsx at the current folder like this:

![image](https://github.com/user-attachments/assets/626d145d-1762-404c-8452-258ab9e66167)

## More practical use
In actual use, you can add other field without excel tag if you don't want to export to excel, like this:

```go

type studentInfo struct {
	Num   int64  `excel:"student number"`
	Name  string `excel:"student name"`
	Birth string `excel:"student birthday"`
	// other field without excel tag if you don't want to export to excel
	other string
}

func main() {
	students := []studentInfo{
		{Num: 1, Name: "John", Birth: "2021-09-21", other: "other"},
		{Num: 2, Name: "Jane", Birth: "2021-11-02", other: "other"},
		{Num: 3, Name: "Jim", Birth: "2022-01-03", other: "other"},
		{Num: 4, Name: "Jill", Birth: "2022-02-01", other: "other"},
		{Num: 5, Name: "Jack", Birth: "2022-04-15", other: "other"},
		{Num: 6, Name: "Jill", Birth: "2022-05-25", other: "other"},
	}
	// no path represent current directory
	err := exceltag.CreateExcel(students, "", "studentInfo.xlsx")
	if err != nil {
		fmt.Println("err", err)
	}
}
```

and the simple spreadsheet of student data will be generated as follows:

![image](https://github.com/user-attachments/assets/bffe6519-51f6-45bb-bed2-f4b7003f090e)

## Autofit column width

The function `AutofitColumn` will automatically fit the column width by the cell content, you can use the `AutofitColumn` function to do this.

```go
err := exceltag.AutofitColumn(f, "Sheet1")
if err != nil {
	fmt.Println("err", err)
}
```

# Resource
[qax-os/excelize
](https://github.com/qax-os/excelize)
