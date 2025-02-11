# exceltag
Generate a simple excel file with tags in your struct

This package will generate a simple excel by your array of any struct with adding tag `excel:"your header"`

For example:
```go
type example struct {
	FieldName  any  `excel:"header"`
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

# Resource
[qax-os/excelize
](https://github.com/qax-os/excelize)
