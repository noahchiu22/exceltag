package exceltag

import (
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
)

func TestCreateExcel(t *testing.T) {
	type studentInfo struct {
		Num       int       `excel:"Student ID"`
		Name      string    `excel:"Name"`
		UpdatedAt time.Time `excel:"Updated time"`
		other     string
	}
	students := []studentInfo{
		{Num: 1, Name: "John", UpdatedAt: time.Now(), other: "other"},
		{Num: 2, Name: "Jane", UpdatedAt: time.Now(), other: "other"},
		{Num: 3, Name: "Jim", UpdatedAt: time.Now(), other: "other"},
		{Num: 4, Name: "Jill", UpdatedAt: time.Now(), other: "other"},
	}
	f, err := CreateExcel(students)
	assert.NoError(t, err)
	assert.NotNil(t, f)

	err = f.SetColWidth("Sheet1", "A", "B", 20)
	assert.NoError(t, err)

	// Save spreadsheet by the given path.
	if err := f.SaveAs("studentInfo.xlsx"); err != nil {
		t.Fatal(err)
	}
}

// func TestAutofitColumn(t *testing.T) {

// }
