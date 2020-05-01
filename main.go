package main

import (
	"fmt"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func main() {
	format := `
	{
		"x_scale": 0.2,
		"y_scale": 0.2,
		"print_obj": true,
		"lock_aspect_ratio": false,
		"locked": true,
		"positioning": "oneCell"
	}`

	f := excelize.NewFile()
	// Insert a picture.
	if err := f.AddPicture("Sheet1", "D5", "image.png", format); err != nil {
		fmt.Println(err)
	}
	f.SetColWidth("Sheet1", "D", "D", 20)
	style1, err := f.NewStyle(`{"fill":{"type":"pattern","color":["#0040FF"],"pattern":1}, "font": {"color":"#777777", "family":"Times New Roman"}}`)
	if err != nil {
		println(err.Error())
	}
	style2, err := f.NewStyle(`{"font": {"color":"#777777", "family":"Sarai"}}`)
	if err != nil {
		println(err.Error())
	}

	f.SetCellValue("Sheet1", "E1", 1)
	f.SetCellValue("Sheet1", "E2", 2)
	f.SetCellValue("Sheet1", "E3", 3)

	// DropDown
	dvRange := excelize.NewDataValidation(true)
	dvRange.Sqref = "A4:A4"
	dvRange.SetSqrefDropList("$E$1:$E$3", true)
	f.AddDataValidation("Sheet1", dvRange)

	f.SetCellValue("Sheet1", "A6", "some text")

	f.SetCellStyle("Sheet1", "A1", "A2", style1)

	f.SetCellStyle("Sheet1", "B1", "B2", style2)

	// Save xlsx file by the given path.
	if err := f.SaveAs("Book1.xlsx"); err != nil {
		fmt.Println(err)
	}
}
